/**
 * Typed tool-failure categories for distinguishing validation, timeout,
 * cancellation, repetition, Office/host, network/provider, protocol, and
 * unexpected internal failures.
 *
 * These replace the legacy `{ok: false, error: "…"}` strings with
 * structured results that carry category and detail for debugging.
 */

export const enum ToolErrorCategory {
  /** Tool name not recognized or arguments failed schema validation (pre-execution). */
  Validation = "validation",
  /** The model call or turn exceeded its absolute time budget. */
  Timeout = "timeout",
  /** The user stopped the agent (AbortError / DOMException). */
  Cancellation = "cancellation",
  /** An identical tool call was repeated too many times in one turn. */
  Repetition = "repetition",
  /** Office.js host or document error (API not supported, etc.). */
  OfficeHost = "office_host",
  /** Network or HTTP error reaching the model provider. */
  Network = "network",
  /** Provider returned an error (4xx/5xx) or malformed response. */
  Provider = "provider",
  /** Protocol-level mismatch (malformed SSE, unexpected finish reason, etc.). */
  Protocol = "protocol",
  /** Unexpected internal error with no better category. */
  Internal = "internal",
}

export interface ToolError {
  ok: false;
  error: string;
  /** Stable machine-readable category. */
  category: ToolErrorCategory;
  /** Provider HTTP status code, if applicable. */
  statusCode?: number;
  /** Original exception name, if available. */
  name?: string;
}

export interface ToolSuccess {
  ok: true;
  [key: string]: unknown;
}

export type ToolResult = ToolSuccess | ToolError;

/** Factory helpers with pre-set categories. */

export function validationError(message: string, name?: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Validation, name };
}

export function timeoutError(message: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Timeout };
}

export function cancellationError(message: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Cancellation };
}

export function repetitionError(message: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Repetition };
}

export function officeHostError(message: string, name?: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.OfficeHost, name };
}

export function networkError(message: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Network };
}

export function providerError(message: string, statusCode?: number): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Provider, statusCode };
}

export function protocolError(message: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Protocol };
}

export function internalError(message: string, name?: string): ToolError {
  return { ok: false, error: message, category: ToolErrorCategory.Internal, name };
}

/**
 * Categorize an unknown error into the best ToolErrorCategory.
 * Used when wrapping exceptions from Office.js or network layers.
 */
export function categorizeError(error: unknown): ToolErrorCategory {
  if (error instanceof DOMException) {
    if (error.name === "TimeoutError") {
      return ToolErrorCategory.Timeout;
    }
    if (error.name === "AbortError") {
      return ToolErrorCategory.Cancellation;
    }
  }

  if (error instanceof TypeError) {
    // fetch() throws TypeError on network failures (CORS, DNS, offline)
    return ToolErrorCategory.Network;
  }

  if (typeof error === "object" && error !== null) {
    const record = error as Record<string, unknown>;
    if (typeof record.category === "string" && [
      "validation", "timeout", "cancellation", "repetition", "office_host", "network", "provider", "protocol", "internal"
    ].includes(record.category)) {
      return record.category as ToolErrorCategory;
    }
    const name = typeof record.name === "string" ? record.name : "";
    const message = typeof record.message === "string" ? record.message : "";

    // Check for attached HTTP status code — always provider error
    if (typeof record.statusCode === "number" || typeof record.status === "number") {
      return ToolErrorCategory.Provider;
    }

    if (name === "TimeoutError") {
      return ToolErrorCategory.Timeout;
    }
    if (name === "AbortError") {
      return ToolErrorCategory.Cancellation;
    }

    // Office.js errors carry a code property
    const code = typeof record.code === "string" ? record.code : "";
    if (code === "ItemNotFound" || message.includes("ItemNotFound")) {
      return ToolErrorCategory.OfficeHost;
    }
    if (
      message.includes("already exists") ||
      message.includes("name conflict") ||
      message.includes("Conflict")
    ) {
      return ToolErrorCategory.OfficeHost;
    }

    // Word-specific: check for WordApi requirement errors
    if (
      name === "OfficeExtension.Error" ||
      message.includes("WordApi") ||
      message.includes("not support")
    ) {
      return ToolErrorCategory.OfficeHost;
    }
  }

  return ToolErrorCategory.Internal;
}

/**
 * Determine if an error is retryable.
 * Retryable: network errors, 429 rate limits, and 5xx server errors.
 * Non-retryable: validation, cancellation, 4xx client errors (except 429),
 * and timeouts when the turn budget is exhausted.
 */
export function isRetryableError(err: ToolError, remainingTurnMs: number): boolean {
  if (err.category === ToolErrorCategory.Cancellation) return false;
  if (err.category === ToolErrorCategory.Validation) return false;
  if (err.category === ToolErrorCategory.Repetition) return false;
  if (err.category === ToolErrorCategory.Timeout) {
    return false;
  }
  if (err.category === ToolErrorCategory.Network) return true;
  if (err.category === ToolErrorCategory.Provider) {
    // Retry on 429 and 5xx, not on 4xx
    if (err.statusCode === 429) return true;
    if (err.statusCode && err.statusCode >= 500 && err.statusCode < 600) return true;
    return false;
  }
  // Protocol, Office, and internal errors are deterministic until fixed.
  return false;
}

/**
 * Sanitize a raw error body to prevent secrets, stack traces, HTML, or
 * giant bodies from reaching the chat or retry aggregates.
 * Extract only a short JSON error.message or status-safe summary.
 */
export function sanitizeErrorBody(raw: string, maxLen: number = 500): string {
  const bounded = raw.slice(0, 16384); // hard read cap

  // Try to extract a short JSON error message
  try {
    const parsed = JSON.parse(bounded);
    if (parsed && typeof parsed === "object") {
      if (typeof parsed.error === "object" && parsed.error !== null) {
        const errObj = parsed.error as Record<string, unknown>;
        if (typeof errObj.message === "string") {
          return sanitizeMessage(errObj.message, maxLen);
        }
      }
      if (typeof parsed.error === "string") {
        return sanitizeMessage(parsed.error, maxLen);
      }
      if (typeof parsed.message === "string") {
        return sanitizeMessage(parsed.message, maxLen);
      }
    }
  } catch {
    // Not JSON, fall through to plain text truncation
  }

  // Remove HTML, stack traces, and excessive whitespace
  const cleaned = bounded
    .replace(/<[^>]*>/g, "")                          // strip HTML
    .replace(/\n\s*at\s+[^\n]*/g, "")                 // strip stack trace lines
    .replace(/\n{3,}/g, "\n\n")                       // collapse excessive newlines
    .replace(/bearer\s+[A-Za-z0-9._~+/=-]+/gi, "[redacted token]")
    .replace(/(api[_-]?key|token|secret)\s*[:=]\s*[^\s,;]+/gi, "$1=[redacted]")
    .trim();

  return sanitizeMessage(cleaned, maxLen);
}

function sanitizeMessage(message: string, maxLen: number): string {
  return message
    .replace(/<[^>]*>/g, "")
    .replace(/\s*\bat\s+[^\n]*/gi, "")
    .replace(/bearer\s+[A-Za-z0-9._~+/=-]+/gi, "[redacted token]")
    .replace(/(api[_-]?key|token|secret)\s*[:=]\s*[^\s,;]+/gi, "$1=[redacted]")
    .trim()
    .slice(0, maxLen);
}
