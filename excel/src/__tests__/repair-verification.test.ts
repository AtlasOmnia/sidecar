/**
 * Tests for dynamic privacy labeling, signal cleanup, JSON fallback,
 * validated-args execution, tracing completeness, and retry-deadline binding.
 */
import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { classifyPrivacy, PrivacyLevel } from "../privacy";
import { categorizeError, ToolErrorCategory } from "../toolErrors";
import { validateToolCall } from "../toolValidation";
import { deadlineAbortSignal, interruptibleSleep, MODEL_CALL_TIMEOUT_MS, TURN_TIMEOUT_MS, throwIfAborted } from "../agent";

// ---------------------------------------------------------------------------
// Privacy labeling
// ---------------------------------------------------------------------------

describe("classifyPrivacy", () => {
  it("classifies 127.0.0.1 as local_bridge", () => {
    expect(classifyPrivacy("http://127.0.0.1:8644/v1")).toBe("local_bridge");
    expect(classifyPrivacy("http://127.0.0.1:1234/v1")).toBe("local_bridge");
  });

  it("classifies localhost as local_bridge", () => {
    expect(classifyPrivacy("http://localhost:8644/v1")).toBe("local_bridge");
    expect(classifyPrivacy("https://localhost:3000")).toBe("local_bridge");
  });

  it("classifies [::1] as local_bridge", () => {
    expect(classifyPrivacy("http://[::1]:8080")).toBe("local_bridge");
  });

  it("classifies known cloud URLs as cloud", () => {
    expect(classifyPrivacy("https://api.openai.com/v1")).toBe("cloud");
    expect(classifyPrivacy("https://api.anthropic.com/v1")).toBe("cloud");
    expect(classifyPrivacy("https://api.deepseek.com/v1")).toBe("cloud");
    expect(classifyPrivacy("https://openrouter.ai/api/v1")).toBe("cloud");
    expect(classifyPrivacy("https://api.together.xyz/v1")).toBe("cloud");
    expect(classifyPrivacy("https://api.groq.com/v1")).toBe("cloud");
    expect(classifyPrivacy("https://api.mistral.ai/v1")).toBe("cloud");
  });

  it("classifies unknown non-loopback as external", () => {
    expect(classifyPrivacy("http://192.168.1.100:1234/v1")).toBe("external");
    expect(classifyPrivacy("https://my-custom-llm.example.com/v1")).toBe("external");
  });

  it("is case-insensitive", () => {
    expect(classifyPrivacy("HTTP://LOCALHOST:8644")).toBe("local_bridge");
    expect(classifyPrivacy("HTTPS://API.OPENAI.COM/v1")).toBe("cloud");
  });

  it("handles URLs with paths correctly", () => {
    expect(classifyPrivacy("http://127.0.0.1:8644/hermes/v1/chat/completions")).toBe("local_bridge");
  });
});

// ---------------------------------------------------------------------------
// categorizeError timeout vs cancellation
// ---------------------------------------------------------------------------

describe("categorizeError timeout semantics", () => {
  it("DOMException TimeoutError → Timeout", () => {
    expect(categorizeError(new DOMException("t/o", "TimeoutError"))).toBe(
      ToolErrorCategory.Timeout,
    );
  });

  it("DOMException AbortError → Cancellation", () => {
    expect(categorizeError(new DOMException("stop", "AbortError"))).toBe(
      ToolErrorCategory.Cancellation,
    );
  });

  it("object with TimeoutError name → Timeout", () => {
    expect(categorizeError({ name: "TimeoutError", message: "t/o" })).toBe(
      ToolErrorCategory.Timeout,
    );
  });

  it("object with AbortError name → Cancellation", () => {
    expect(categorizeError({ name: "AbortError", message: "stop" })).toBe(
      ToolErrorCategory.Cancellation,
    );
  });
});

// ---------------------------------------------------------------------------
// Deadline signal
// ---------------------------------------------------------------------------

describe("deadlineAbortSignal", () => {
  it("creates an already-aborted signal when ms <= 0", () => {
    const deadline = deadlineAbortSignal(0, "test");
    expect(deadline.signal.aborted).toBe(true);
    deadline.cleanup();
  });

  it("fires after the specified time", async () => {
    const deadline = deadlineAbortSignal(50, "test deadline");
    expect(deadline.signal.aborted).toBe(false);
    await new Promise((resolve) => {
      deadline.signal.addEventListener("abort", () => resolve(undefined));
    });
    expect(deadline.signal.aborted).toBe(true);
    deadline.cleanup();
  });

  it("cleanup prevents the timer from firing", async () => {
    const deadline = deadlineAbortSignal(50, "test");
    deadline.cleanup();
    // Wait past the deadline; signal should not abort
    await new Promise((resolve) => setTimeout(resolve, 80));
    expect(deadline.signal.aborted).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Interruptible sleep
// ---------------------------------------------------------------------------

describe("interruptibleSleep", () => {
  it("resolves after the delay when not aborted", async () => {
    const start = Date.now();
    await interruptibleSleep(50, new AbortController().signal);
    expect(Date.now() - start).toBeGreaterThanOrEqual(40);
  });

  it("rejects with AbortError when signal fires", async () => {
    const controller = new AbortController();
    const promise = interruptibleSleep(5000, controller.signal);
    controller.abort();

    await expect(promise).rejects.toThrow(DOMException);
    await expect(promise).rejects.toThrow("Aborted during retry delay.");
  });
});

// ---------------------------------------------------------------------------
// Validated args are execution authority
// ---------------------------------------------------------------------------

describe("validated args execution authority", () => {
  it("validateToolCall returns parsed args that can be passed to execution", () => {
    const result = validateToolCall(
      "read_range",
      JSON.stringify({ sheet: "Data", address: "A1:C10" }),
    );
    expect("valid" in result && result.valid).toBe(true);
    if ("valid" in result) {
      // The parsed args are directly usable — no re-parsing needed
      expect(result.parsedArgs.sheet).toBe("Data");
      expect(result.parsedArgs.address).toBe("A1:C10");
    }
  });

  it("validation rejects malformed args that would crash JSON.parse", () => {
    const result = validateToolCall("read_range", "{{bad json");
    expect("category" in result).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
    }
  });

  it("validation rejects unknown tools before any args processing", () => {
    const result = validateToolCall("nonexistent", '{"x":1}');
    expect("category" in result).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
      expect(result.error).toContain("Unknown tool");
    }
  });

  it("validation rejects unsafe payloads that aren't objects", () => {
    const result = validateToolCall("read_range", '"just a string"');
    expect("category" in result).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
    }
  });
});

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

describe("timeout constants", () => {
  it("MODEL_CALL_TIMEOUT_MS is 20s", () => {
    expect(MODEL_CALL_TIMEOUT_MS).toBe(20_000);
  });

  it("TURN_TIMEOUT_MS is 90s", () => {
    expect(TURN_TIMEOUT_MS).toBe(90_000);
  });
});

// ---------------------------------------------------------------------------
// throwIfAborted
// ---------------------------------------------------------------------------

describe("throwIfAborted", () => {
  it("throws AbortError when signal is aborted", () => {
    const controller = new AbortController();
    controller.abort();
    expect(() => throwIfAborted(controller.signal)).toThrow(DOMException);
  });

  it("does not throw when signal is not aborted", () => {
    const controller = new AbortController();
    expect(() => throwIfAborted(controller.signal)).not.toThrow();
  });
});
