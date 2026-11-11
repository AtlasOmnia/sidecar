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
  /** Office.js host or workbook error (sheet not found, API not supported, etc.). */
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

    // Error-like with agent-specific names
    if (name === "ExcelToolError" || name === "InvalidArgument") {
      return ToolErrorCategory.OfficeHost;
    }
  }

  return ToolErrorCategory.Internal;
}
