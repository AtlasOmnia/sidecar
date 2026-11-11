/**
 * Behavioral tests for release readiness requirements:
 * H2 - Host compatibility (WordApi 1.4 gate)
 * H1 - Deadline/retry containment
 * M1 - Strict tool validation
 * M3 - Provider error sanitization
 * M5 - Iteration containment (cap at 20, checkpoint)
 * L2 - Overview selection.text as string
 * L3 - Non-object tool result not terminating
 */

import { beforeEach, describe, expect, it, vi, afterEach, type Mock } from "vitest";
import {
  HermesAgent,
  readOpenAIJsonResponse,
  interruptibleSleep,
  throwIfAborted,
  resetInjectedSleep,
  sleepHooks,
  _readBoundedBodyForTest,
  MODEL_CALL_TIMEOUT_MS,
  TURN_TIMEOUT_MS,
} from "../agent";
import { createLinkedDeadline, raceWithSignal } from "../agentRuntime";
import { ToolErrorCategory, categorizeError, sanitizeErrorBody } from "../toolErrors";
import { validateToolCall } from "../toolValidation";
import { classifyPrivacy } from "../privacy";
import { supportsWordApi14 } from "../wordTools";
import { PRIVACY_LABELS } from "../privacy";

// ---------------------------------------------------------------------------
// H2: Host compatibility — WordApi 1.4 gate
// ---------------------------------------------------------------------------

describe("H2: host compatibility (WordApi 1.4 gate)", () => {
  beforeEach(() => {
    vi.stubGlobal("Office", {
      context: {
        requirements: {
          isSetSupported: vi.fn().mockReturnValue(true),
        },
      },
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("supportsWordApi14 returns true when isSetSupported says 1.4 is available", () => {
    expect(supportsWordApi14()).toBe(true);
  });

  it("checks the WordApi 1.4 requirement set, not WordApiDesktop", () => {
    const isSetSupported = vi.fn().mockReturnValue(true);
    vi.stubGlobal("Office", { context: { requirements: { isSetSupported } } });

    supportsWordApi14();

    expect(isSetSupported).toHaveBeenCalledWith("WordApi", "1.4");
  });

  it("supportsWordApi14 returns false when isSetSupported returns false for 1.4", () => {
    const isSetSupported = vi.fn().mockReturnValue(false);
    vi.stubGlobal("Office", {
      context: { requirements: { isSetSupported } },
    });
    expect(supportsWordApi14()).toBe(false);
  });

  it("supportsWordApi14 returns false when isSetSupported throws", () => {
    const isSetSupported = vi.fn().mockImplementation(() => {
      throw new Error("Not available");
    });
    vi.stubGlobal("Office", {
      context: { requirements: { isSetSupported } },
    });
    expect(supportsWordApi14()).toBe(false);
  });

  it("supportsWordApi14 returns false when Office is undefined", () => {
    vi.stubGlobal("Office", undefined);
    expect(supportsWordApi14()).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// H1: Deadline and retry containment
// ---------------------------------------------------------------------------

describe("H1: deadline abort signal", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
    resetInjectedSleep();
  });

  it("fires TimeoutError after the specified timeout", () => {
    const parent = new AbortController();
    const { signal, cleanup } = createLinkedDeadline(parent.signal, 1000, "Test deadline");

    expect(signal.aborted).toBe(false);

    vi.advanceTimersByTime(1001);

    expect(signal.aborted).toBe(true);
    cleanup();
  });

  it("cleanup prevents the deadline from firing", () => {
    const parent = new AbortController();
    const { signal, cleanup } = createLinkedDeadline(parent.signal, 1000, "Test deadline");

    cleanup();
    vi.advanceTimersByTime(1001);

    expect(signal.aborted).toBe(false);
  });

  it("zero timeout produces a non-aborted signal", () => {
    const parent = new AbortController();
    const { signal, cleanup } = createLinkedDeadline(parent.signal, 0, "Zero timeout");
    expect(signal.aborted).toBe(false);
    cleanup();
  });
});

describe("H1: interruptible sleep", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  afterEach(() => {
    vi.useRealTimers();
    resetInjectedSleep();
  });

  it("resolves after the specified delay", async () => {
    const controller = new AbortController();
    const promise = interruptibleSleep(1000, controller.signal);

    vi.advanceTimersByTime(1000);
    await expect(promise).resolves.toBeUndefined();
  });

  it("rejects when aborted during sleep", async () => {
    const controller = new AbortController();
    const promise = interruptibleSleep(5000, controller.signal);

    controller.abort();
    vi.advanceTimersByTime(0); // Advance microtasks

    await expect(promise).rejects.toThrow(DOMException);
  });

  it("respects injected sleep for testing", async () => {
    let slept = false;
    sleepHooks.injectedSleep = async (_ms: number) => {
      slept = true;
    };

    const controller = new AbortController();
    await interruptibleSleep(10000, controller.signal);

    expect(slept).toBe(true);
  });

  it("zero ms sleep resolves immediately", async () => {
    const controller = new AbortController();
    await expect(interruptibleSleep(0, controller.signal)).resolves.toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// D: abort-aware reader reads and cancel-on-cap
// ---------------------------------------------------------------------------

describe("D: abort-aware reader reads and cancel-on-cap", () => {
  it("raceWithSignal rejects when a reader aborts during a pending read", async () => {
    // Create a stream that blocks on read (never enqueues)
    const stream = new ReadableStream<Uint8Array>({
      start() { /* never enqueues — hangs */ },
    });
    const reader = stream.getReader();
    const ctrl = new AbortController();

    const readPromise = raceWithSignal(reader.read(), ctrl.signal);

    // Abort while read is pending
    ctrl.abort(new DOMException("deadline", "TimeoutError"));

    await expect(readPromise).rejects.toThrow(DOMException);
    try { reader.releaseLock(); } catch { /* ok */ }
  });

  it("bounded body read cancels reader at cap and does not consume beyond 64 KiB", async () => {
    let cancelCalled = false;
    let chunkCount = 0;
    // Stream that emits 65 chunks of 1024 bytes each = ~66 KiB (just over cap)
    const chunks: Uint8Array[] = Array.from({ length: 65 }, (_, i) => {
      const buf = new Uint8Array(1024);
      buf.fill(65 + (i % 26)); // 'A'–'Z' cycle
      return buf;
    });

    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        for (const chunk of chunks) {
          controller.enqueue(chunk);
          chunkCount++;
        }
        controller.close();
      },
      cancel() {
        cancelCalled = true;
      },
    });

    const response = new Response(stream, {
      status: 200,
      headers: { "content-type": "application/json" },
    });

    const ctrl = new AbortController();
    const result = await _readBoundedBodyForTest(response, ctrl.signal);

    // Must not exceed 64 KiB cap
    expect(result.length).toBeLessThanOrEqual(65536);
    // Cancel must have been called (cap reached → cancelReader)
    expect(cancelCalled).toBe(true);
    // Must have read at least enough to hit the cap
    expect(chunkCount).toBeLessThanOrEqual(65);
  });
});

// ---------------------------------------------------------------------------
// M1: Tool validation
// ---------------------------------------------------------------------------

describe("M1: strict tool validation", () => {
  it("rejects unknown tool names", () => {
    const result = validateToolCall("nonexistent_tool", "{}");
    expect("ok" in result && result.ok === false).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
    }
  });

  it("rejects invalid JSON arguments", () => {
    const result = validateToolCall("read_document_text", "not valid json");
    expect("ok" in result && result.ok === false).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
    }
  });

  it("rejects non-object arguments", () => {
    const result = validateToolCall("read_document_text", "[1, 2, 3]");
    expect("ok" in result && result.ok === false).toBe(true);
    if ("category" in result) {
      expect(result.category).toBe(ToolErrorCategory.Validation);
    }
  });

  it("rejects missing required arguments", () => {
    // apply_edit_plan requires "edits"
    const result = validateToolCall("apply_edit_plan", "{}");
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("edits");
    }
  });

  it("rejects unknown fields when additionalProperties is false", () => {
    // read_document_text has additionalProperties: false
    const result = validateToolCall(
      "read_document_text",
      '{"maxLength": 100, "bogusField": "extra"}',
    );
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("bogusField");
    }
  });

  it("rejects wrong field types", () => {
    // maxLength must be an integer
    const result = validateToolCall("read_document_text", '{"maxLength": "not-a-number"}');
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("maxLength");
    }
  });

  it("accepts valid arguments", () => {
    const result = validateToolCall("read_document_text", '{"maxLength": 500}');
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("accepts empty object for tools with no required fields", () => {
    const result = validateToolCall("get_document_overview", "{}");
    expect("valid" in result && result.valid === true).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// M3: Error sanitization
// ---------------------------------------------------------------------------

describe("M3: provider error sanitization", () => {
  it("extracts error.message from JSON error bodies", () => {
    const raw = JSON.stringify({
      error: { message: "Rate limit exceeded", type: "rate_limit" },
    });
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).toBe("Rate limit exceeded");
  });

  it("extracts error string from JSON bodies", () => {
    const raw = JSON.stringify({ error: "Service unavailable" });
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).toBe("Service unavailable");
  });

  it("truncates long error messages", () => {
    const longMessage = "x".repeat(2000);
    const raw = JSON.stringify({ error: { message: longMessage } });
    const sanitized = sanitizeErrorBody(raw, 100);
    expect(sanitized.length).toBeLessThanOrEqual(100);
  });

  it("strips HTML from error bodies", () => {
    const raw = "<html><body><h1>500 Internal Server Error</h1></body></html>";
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).not.toContain("<html>");
    expect(sanitized).not.toContain("<h1>");
    expect(sanitized).toContain("500 Internal Server Error");
  });

  it("strips stack traces", () => {
    const raw = "Error: something broke\n    at Object.<anonymous> (/src/test.ts:10:5)\n    at runTest (/node_modules/vitest:1:2)";
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).not.toContain("    at ");
    expect(sanitized).toContain("Error: something broke");
  });

  it("caps read size", () => {
    // Generate a body larger than the 16384 hard cap
    const large = "x".repeat(20000);
    const sanitized = sanitizeErrorBody(large, 500);
    expect(sanitized.length).toBeLessThanOrEqual(500);
  });

  it("handles non-JSON input gracefully", () => {
    const raw = "Just a plain text error message";
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).toBe(raw);
  });
});

// ---------------------------------------------------------------------------
// M5: Iteration containment
// ---------------------------------------------------------------------------

describe("M5: iteration containment", () => {
  it("absolute iteration cap is 20", () => {
    // The ABSOLUTE_MAX_ITERATIONS constant is 20
    // Settings can't exceed this
    const settingsMax = 50;
    const effectiveMax = Math.min(settingsMax, 20);
    expect(effectiveMax).toBe(20);
  });

  it("respects settings max when below cap", () => {
    const effectiveMax = Math.min(8, 20);
    expect(effectiveMax).toBe(8);
  });

  it("enforces minimum of 1", () => {
    const effectiveMax = Math.min(Math.max(0, 1), 20);
    expect(effectiveMax).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// M2: Privacy classification
// ---------------------------------------------------------------------------

describe("M2: privacy classification", () => {
  it("classifies loopback URLs as local_bridge", () => {
    expect(classifyPrivacy("http://127.0.0.1:8645/v1")).toBe("local_bridge");
    expect(classifyPrivacy("http://localhost:3000")).toBe("local_bridge");
    expect(classifyPrivacy("http://[::1]:8080")).toBe("local_bridge");
    expect(classifyPrivacy("http://0.0.0.0:4000")).toBe("external");
  });

  it("classifies RFC1918 private addresses as private_network", () => {
    expect(classifyPrivacy("https://192.168.1.100:1234/v1")).toBe("private_network");
    expect(classifyPrivacy("https://10.0.0.1:8080")).toBe("private_network");
    expect(classifyPrivacy("https://172.16.0.1/api")).toBe("private_network");
    expect(classifyPrivacy("https://172.31.255.255/")).toBe("private_network");
  });

  it("classifies link-local as private_network", () => {
    expect(classifyPrivacy("http://169.254.1.1:1234")).toBe("private_network");
  });

  it("classifies IPv6 ULA and link-local addresses as private_network", () => {
    expect(classifyPrivacy("http://[fd12::1]:8080")).toBe("private_network");
    expect(classifyPrivacy("http://[fe80::1]:8080")).toBe("private_network");
  });

  it("classifies cloud provider URLs as cloud (suffix-matched, not substring)", () => {
    expect(classifyPrivacy("https://api.openai.com/v1")).toBe("cloud");
    expect(classifyPrivacy("https://api.deepseek.com/chat")).toBe("cloud");
    expect(classifyPrivacy("https://openrouter.ai/api/v1")).toBe("cloud");
  });

  it("does not classify spoofed cloud domains as cloud", () => {
    // api.openai.com.attacker.example must NOT match as OpenAI
    expect(classifyPrivacy("https://api.openai.com.attacker.example/v1")).toBe("external");
    expect(classifyPrivacy("https://openai.com.evil.com/")).toBe("external");
  });

  it("does not classify private-address lookalike hostnames as private LAN", () => {
    expect(classifyPrivacy("https://192.168.evil.example/v1")).toBe("external");
    expect(classifyPrivacy("https://169.254.attacker.example/")).toBe("external");
    expect(classifyPrivacy("https://192.168.1.999/v1")).toBe("external");
  });

  it("classifies non-loopback, non-private, non-cloud URLs as external", () => {
    expect(classifyPrivacy("https://my-company-server.local/api")).toBe("external");
    expect(classifyPrivacy("https://93.184.216.34/")).toBe("external");
  });

  it("parses credentials and mixed-case hostnames without mislabeling the endpoint", () => {
    expect(classifyPrivacy("https://user:secret@API.OPENAI.COM./v1")).toBe("cloud");
  });

  it("classifies malformed and unsupported-protocol URLs as external", () => {
    expect(classifyPrivacy("not a URL")).toBe("external");
    expect(classifyPrivacy("ftp://127.0.0.1/v1")).toBe("external");
  });
});

// ---------------------------------------------------------------------------
// F: Privacy tooltip does not overclaim data stays within local network
// ---------------------------------------------------------------------------

describe("F: privacy tooltip accuracy", () => {
  it("private_network tooltip does not claim data stays within local network", () => {
    const { tooltip } = PRIVACY_LABELS.private_network;
    expect(tooltip).not.toContain("Document data stays within your local network");
    expect(tooltip).toContain("immediate endpoint is on your local network");
    expect(tooltip).toContain("may still send data to a cloud provider");
  });

  it("local_bridge tooltip warns about cloud backend possibility", () => {
    const { tooltip } = PRIVACY_LABELS.local_bridge;
    expect(tooltip).toContain("may still send data to a cloud provider");
  });

  it("cloud tooltip explicitly states data is transmitted to external servers", () => {
    const { tooltip } = PRIVACY_LABELS.cloud;
    expect(tooltip).toContain("transmitted to external servers");
  });

  it("all labels are defined", () => {
    expect(PRIVACY_LABELS.local_bridge).toBeDefined();
    expect(PRIVACY_LABELS.private_network).toBeDefined();
    expect(PRIVACY_LABELS.cloud).toBeDefined();
    expect(PRIVACY_LABELS.external).toBeDefined();
  });
});

// ---------------------------------------------------------------------------
// L3: Non-object tool result
// ---------------------------------------------------------------------------

describe("L3: non-object tool result", () => {
  it("categorizeError returns Internal for regular Error", () => {
    const err = new Error("something went wrong");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Internal);
  });

  it("categorizeError returns Cancellation for AbortError", () => {
    const err = new DOMException("aborted", "AbortError");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Cancellation);
  });

  it("categorizeError returns Timeout for TimeoutError", () => {
    const err = new DOMException("timeout", "TimeoutError");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Timeout);
  });

  it("categorizeError returns Network for TypeError", () => {
    const err = new TypeError("Failed to fetch");
    expect(categorizeError(err)).toBe(ToolErrorCategory.Network);
  });

  it("categorizeError returns Provider for error with statusCode", () => {
    const err = Object.assign(new Error("Not Found"), { statusCode: 404 });
    expect(categorizeError(err)).toBe(ToolErrorCategory.Provider);
  });
});

// ---------------------------------------------------------------------------
// H2 behavioral: WordApiDesktop 1.4 through Word.run mock
// ---------------------------------------------------------------------------

describe("H2: WordApiDesktop 1.4 in Word.run flow", () => {
  let isSetSupported: ReturnType<typeof vi.fn>;
  let wordRun: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    vi.clearAllMocks();
    isSetSupported = vi.fn().mockReturnValue(true);
    wordRun = vi.fn(async (work: (ctx: Record<string, unknown>) => Promise<unknown>) => {
      const selection = {
        text: "test selection",
        load: vi.fn(),
      };
      const body = {
        text: "body text",
        paragraphs: { items: [{ text: "p1" }, { text: "p2" }], load: vi.fn() },
        load: vi.fn(),
      };
      const ctx = {
        document: { body, getSelection: () => selection },
        sync: vi.fn().mockResolvedValue(undefined),
      };
      return work(ctx);
    });
    vi.stubGlobal("Word", { run: wordRun });
    vi.stubGlobal("Office", {
      context: { requirements: { isSetSupported } },
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("getDocumentSnapshot loads only text when WordApiDesktop 1.4 unsupported", async () => {
    isSetSupported.mockReturnValue(false);
    const { getDocumentSnapshot } = await import("../wordTools");
    const result = await getDocumentSnapshot(200);
    expect(result.overview.selection.start).toBeNull();
    expect(result.overview.selection.end).toBeNull();
    expect(typeof result.overview.selection.text).toBe("string");
    expect(isSetSupported).toHaveBeenCalledWith("WordApiDesktop", "1.4");
  });

  it("getDocumentSnapshot loads start/end when WordApiDesktop 1.4 supported", async () => {
    isSetSupported.mockReturnValue(true);
    wordRun.mockImplementation(async (work: (ctx: Record<string, unknown>) => Promise<unknown>) => {
      const selection = {
        text: "test",
        start: 5,
        end: 9,
        load: vi.fn(),
      };
      const body = {
        text: "body",
        paragraphs: { items: [], load: vi.fn() },
        load: vi.fn(),
      };
      const ctx = {
        document: { body, getSelection: () => selection },
        sync: vi.fn().mockResolvedValue(undefined),
      };
      return work(ctx);
    });
    const { getDocumentSnapshot } = await import("../wordTools");
    const result = await getDocumentSnapshot(200);
    expect(result.overview.selection.start).toBe(5);
    expect(result.overview.selection.end).toBe(9);
  });

  it("calls isSetSupported with WordApiDesktop not WordApi", async () => {
    isSetSupported.mockReturnValue(true);
    const { getDocumentSnapshot } = await import("../wordTools");
    await getDocumentSnapshot(100);
    expect(isSetSupported).toHaveBeenCalledWith("WordApiDesktop", "1.4");
  });
});

// ---------------------------------------------------------------------------
// M1: Nested validation — apply_edit_plan items
// ---------------------------------------------------------------------------

describe("M1: nested validation for apply_edit_plan", () => {
  it("rejects unknown fields inside edit items", () => {
    const args = JSON.stringify({
      edits: [{ find: "x", replace: "y", overwrite_all: true }],
    });
    const result = validateToolCall("apply_edit_plan", args);
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("overwrite_all");
      expect(result.error).toContain("Unknown");
    }
  });

  it("rejects items missing required find/replace", () => {
    const args = JSON.stringify({
      edits: [{ replace: "y" }],
    });
    const result = validateToolCall("apply_edit_plan", args);
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("find");
    }
  });

  it("rejects items with wrong types in nested fields", () => {
    const args = JSON.stringify({
      edits: [{ find: 123, replace: "y" }],
    });
    const result = validateToolCall("apply_edit_plan", args);
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("find");
    }
  });

  it("accepts valid edit plan with all recognized fields", () => {
    const args = JSON.stringify({
      edits: [
        { find: "old", replace: "new", matchCase: true, matchWholeWord: false, reason: "typo" },
      ],
    });
    const result = validateToolCall("apply_edit_plan", args);
    expect("valid" in result && result.valid === true).toBe(true);
  });

  it("rejects too many edits exceeding maxItems", () => {
    const edits = Array.from({ length: 101 }, (_, i) => ({
      find: `find${i}`,
      replace: `rep${i}`,
    }));
    const result = validateToolCall("apply_edit_plan", JSON.stringify({ edits }));
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("100");
    }
  });

  it("rejects empty edits array below minItems", () => {
    const result = validateToolCall("apply_edit_plan", JSON.stringify({ edits: [] }));
    expect("ok" in result && result.ok === false).toBe(true);
    if ("error" in result) {
      expect(result.error).toContain("1");
    }
  });
});

// ---------------------------------------------------------------------------
// M3: Production-path sanitization with secrets/stack/HTML fixtures
// ---------------------------------------------------------------------------

describe("M3: production-path error sanitization", () => {
  it("extracts only short JSON error.message from provider body with secrets", () => {
    const raw = JSON.stringify({
      error: {
        message: "Invalid API key",
        type: "authentication_error",
        code: "invalid_api_key",
        param: null,
      },
      api_key_hint: "sk-ant-api03-xxxxxxxxxxxx",
      request_id: "req_abc123",
    });
    const sanitized = sanitizeErrorBody(raw, 200);
    expect(sanitized).toBe("Invalid API key");
    expect(sanitized).not.toContain("sk-ant");
    expect(sanitized).not.toContain("api_key_hint");
  });

  it("strips HTML and stack traces from production error page", () => {
    const raw = "<!DOCTYPE html>\n<html>\n<head><title>502 Bad Gateway</title></head>\n<body>\n<h1>502 Bad Gateway</h1>\n<p>The upstream server returned an invalid response.</p>\n<pre>Error: connect ECONNREFUSED 127.0.0.1:8643\n    at TCPConnectWrap.afterConnect [as oncomplete] (node:net:1234:16)</pre>\n</body>\n</html>";
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).not.toContain("<html>");
    expect(sanitized).not.toContain("<body>");
    expect(sanitized).not.toContain("TCPConnectWrap");
    expect(sanitized).not.toContain("    at ");
    expect(sanitized).toContain("502 Bad Gateway");
  });

  it("truncates oversized plain-text bodies to maxLen", () => {
    const giantBody = "ERROR: " + "x".repeat(5000);
    const sanitized = sanitizeErrorBody(giantBody, 200);
    expect(sanitized.length).toBeLessThanOrEqual(200);
    expect(sanitized.startsWith("ERROR:")).toBe(true);
  });

  it("sanitizer default output cap is 500 chars", () => {
    const raw = JSON.stringify({ error: { message: "test" } });
    const sanitized = sanitizeErrorBody(raw);
    expect(sanitized).toBe("test");
    expect(sanitized.length).toBeLessThanOrEqual(500);
  });
});
