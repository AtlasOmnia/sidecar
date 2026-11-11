export enum ToolErrorCategory {
  Validation = "validation", Timeout = "timeout", Cancellation = "cancellation", Repetition = "repetition", OfficeHost = "office_host", Network = "network", Provider = "provider", Protocol = "protocol", Internal = "internal"
}
export type ToolErrorCategoryName = `${ToolErrorCategory}`;
export interface ToolError { ok: false; error: string; category: ToolErrorCategoryName; statusCode?: number; name?: string; details?: Record<string, unknown>; requestId?: string; }
export interface ToolErrorOptions { fetchTypeError?: boolean; }

type TrustedToolError = ToolError & { [trustedToolError]?: true };
const trustedToolError = Symbol("trustedToolError");
const SAFE_ERROR_NAMES = new Set(["Error", "TypeError", "RangeError", "SyntaxError", "ReferenceError", "URIError", "EvalError", "DOMException", "AbortError", "TimeoutError"]);
const SAFE_CODE = /^[A-Za-z0-9]{1,32}$/;
const SAFE_TOKEN = /^[A-Za-z0-9._-]{1,64}$/;
const SAFE_CAPABILITY_TEXT = /^[\x20-\x7e]{1,512}$/;

const FIXED_MESSAGES: Record<ToolErrorCategoryName, string> = {
  validation: "The tool request was invalid.",
  timeout: "The operation exceeded its time limit.",
  cancellation: "The running Hermes task was stopped.",
  repetition: "The repeated tool request was blocked.",
  office_host: "PowerPoint could not complete the requested operation.",
  network: "Could not reach the model endpoint.",
  provider: "The model provider returned an error.",
  protocol: "The model response could not be processed.",
  internal: "An internal Hermes error occurred.",
};

function messageOf(value: unknown): string { if (typeof value === "string") return value; if (value instanceof Error) return value.message || value.name; if (value && typeof value === "object" && typeof (value as Record<string, unknown>).message === "string") return String((value as Record<string, unknown>).message); return "Operation failed."; }
function isCategory(value: unknown): value is ToolErrorCategoryName { return typeof value === "string" && Object.values(ToolErrorCategory).includes(value as ToolErrorCategory); }
function safeName(value: unknown): string | undefined { return typeof value === "string" && SAFE_ERROR_NAMES.has(value) ? value : undefined; }
function safeNumber(value: unknown): number | undefined { return typeof value === "number" && Number.isFinite(value) ? value : undefined; }
function safeCode(value: unknown): string | undefined { return typeof value === "string" && SAFE_CODE.test(value) ? value : undefined; }
function safeToken(value: unknown): string | undefined { return typeof value === "string" && SAFE_TOKEN.test(value) ? value : undefined; }

function sanitizedDetails(record: Record<string, unknown> | undefined, trusted: boolean, allowCapabilityMetadata: boolean): Record<string, unknown> | undefined {
  const source = record?.details && typeof record.details === "object" && !Array.isArray(record.details)
    ? record.details as Record<string, unknown>
    : undefined;
  const result: Record<string, unknown> = {};
  const code = safeCode(source?.code) ?? safeCode(record?.code);
  if (code) result.code = code;
  if (!trusted) {
    if (allowCapabilityMetadata && typeof source?.available === "boolean") result.available = source.available;
    if (allowCapabilityMetadata && typeof source?.requiredApi === "string" && SAFE_CAPABILITY_TEXT.test(source.requiredApi)) result.requiredApi = source.requiredApi;
    return Object.keys(result).length ? result : undefined;
  }
  if (typeof source?.available === "boolean") result.available = source.available;
  if (typeof source?.requiredApi === "string" && SAFE_CAPABILITY_TEXT.test(source.requiredApi)) result.requiredApi = source.requiredApi;
  if (typeof source?.reason === "string" && SAFE_CAPABILITY_TEXT.test(source.reason)) result.reason = source.reason;
  for (const key of ["slideIndex", "textLength", "imageLength", "beforeLength", "afterLength"]) {
    const number = safeNumber(source?.[key]);
    if (number !== undefined) result[key] = number;
  }
  return Object.keys(result).length ? result : undefined;
}

function makeTrustedToolError(category: ToolErrorCategoryName, error: string, extras: Partial<ToolError> = {}): ToolError {
  const result: TrustedToolError = { ok: false, error, category, ...extras };
  Object.defineProperty(result, trustedToolError, { value: true, enumerable: false });
  return result;
}

export function asToolError(value: unknown, fallbackCategory: ToolErrorCategoryName = "internal", options: ToolErrorOptions = {}): ToolError {
  const record = value && typeof value === "object" ? value as Record<string, unknown> : undefined;
  const trusted = Boolean(record && (record as unknown as TrustedToolError)[trustedToolError] === true);
  const hasExplicitCategory = isCategory(record?.category);
  const existingCategory: ToolErrorCategoryName = hasExplicitCategory ? record?.category as ToolErrorCategoryName : record?.code ? "office_host" : fallbackCategory;
  const explicitStatus = safeNumber(record?.statusCode);
  const status = safeNumber(record?.status);
  const statusCode = status ?? explicitStatus;
  const name = safeName(value instanceof Error ? value.name : record?.name);
  if (name === "AbortError") return makeTrustedToolError("cancellation", FIXED_MESSAGES.cancellation, { name });
  if (name === "TimeoutError") return makeTrustedToolError("timeout", FIXED_MESSAGES.timeout, { name });
  if (value instanceof TypeError && (options.fetchTypeError === true || fallbackCategory === "network")) return makeTrustedToolError("network", FIXED_MESSAGES.network, { name: "TypeError" });

  const category = statusCode !== undefined && !trusted && !hasExplicitCategory ? "provider" : existingCategory;
  const details = sanitizedDetails(record, trusted, existingCategory === "office_host");
  const extras: Partial<ToolError> = {
    ...(statusCode === undefined ? {} : { statusCode }),
    ...(name ? { name } : {}),
    ...(details ? { details } : {}),
    ...(safeToken(record?.requestId) ? { requestId: safeToken(record?.requestId) } : {}),
  };
  const trustedMessage = typeof record?.error === "string" ? record.error : messageOf(value);
  const error = trusted ? trustedMessage : FIXED_MESSAGES[category];
  return makeTrustedToolError(category, error, extras);
}

export function stringifyToolError(value: unknown, fallbackCategory?: ToolErrorCategoryName): string { return JSON.stringify(asToolError(value, fallbackCategory)); }
export function errorToToolError(error: unknown): ToolError { return asToolError(error, "internal", { fetchTypeError: false }); }
export function categorizeError(error: unknown): ToolErrorCategoryName { return asToolError(error, "internal", { fetchTypeError: false }).category; }
export const validationError = (message: string, name?: string) => { const safe = safeName(name); return makeTrustedToolError("validation", message, safe ? { name: safe } : {}); };
export const timeoutError = (message: string) => asToolError(new DOMException(message, "TimeoutError"));
export const cancellationError = (message: string) => asToolError(new DOMException(message, "AbortError"));
export const repetitionError = (message: string) => makeTrustedToolError("repetition", message);
export const officeHostError = (message: string) => makeTrustedToolError("office_host", message);
export const networkError = (message: string) => makeTrustedToolError("network", message);
export const providerError = (message: string, statusCode?: number) => makeTrustedToolError("provider", message, statusCode === undefined ? {} : { statusCode });
export const protocolError = (message: string) => makeTrustedToolError("protocol", message);
export const internalError = (message: string) => makeTrustedToolError("internal", message);
export interface ToolResult { ok: boolean; [key: string]: unknown; }
