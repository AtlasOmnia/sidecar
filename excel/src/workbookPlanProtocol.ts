import type { WorkbookPlan } from "./generated/workbookPlanSchema";
import { canonicalizeWorkbookPlan } from "./workbookPlanCanonicalization";

/** The only plan tool name admitted by the V1 protocol. */
export const WORKBOOK_PLAN_TOOL_NAME = "excel_execute_plan" as const;

/** Conservative protocol bounds applied before JSON/XML parsing. */
export const MAX_PROTOCOL_TOOL_CALLS = 16;
export const MAX_PROTOCOL_CONTENT_BYTES = 128 * 1024;
export const MAX_PROTOCOL_ARGUMENT_BYTES = 64 * 1024;
/**
 * The serialized envelope must fit every bounded primitive call, plus a
 * deterministic amount of per-call JSON framing overhead.
 */
export const MAX_PROTOCOL_TOOL_CALL_BYTES =
  MAX_PROTOCOL_TOOL_CALLS * (MAX_PROTOCOL_ARGUMENT_BYTES + 1024);

const PROTOCOL_REJECTION = Object.freeze({
  code: "protocol_rejected" as const,
  message: "Invalid workbook plan protocol envelope" as const,
});

export interface WorkbookPlanProtocolError {
  readonly code: "protocol_rejected";
  readonly message: "Invalid workbook plan protocol envelope";
}

export interface NormalizedWorkbookPlan {
  readonly ok: true;
  readonly kind: "plan";
  readonly transport: "openai-tool-call" | "xml-fallback";
  readonly plan: WorkbookPlan;
  readonly semanticDigest: string;
}

export interface NoWorkbookPlan {
  readonly ok: true;
  readonly kind: "none";
}

export interface RejectedWorkbookPlanProtocol {
  readonly ok: false;
  readonly error: WorkbookPlanProtocolError;
}

export type WorkbookPlanProtocolResult =
  | NormalizedWorkbookPlan
  | NoWorkbookPlan
  | RejectedWorkbookPlanProtocol;

type PlainRecord = Record<string, unknown>;

type OpenAIPlanCandidate = {
  readonly payload: unknown;
};

type ArrayElementRead =
  | { readonly ok: true; readonly value: unknown }
  | { readonly ok: false };

function reject(): RejectedWorkbookPlanProtocol {
  return Object.freeze({
    ok: false as const,
    error: Object.freeze({ ...PROTOCOL_REJECTION }),
  });
}

function none(): NoWorkbookPlan {
  return Object.freeze({ ok: true as const, kind: "none" as const });
}

function isPlainRecord(value: unknown): value is PlainRecord {
  try {
    if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
    return Object.getPrototypeOf(value) === Object.prototype;
  } catch {
    return false;
  }
}

function ownDataDescriptor(value: object, key: PropertyKey): PropertyDescriptor | undefined {
  try {
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (descriptor === undefined || !Object.prototype.hasOwnProperty.call(descriptor, "value")) {
      return undefined;
    }
    return descriptor;
  } catch {
    return undefined;
  }
}

function ownDataValue(value: object, key: PropertyKey): unknown {
  return ownDataDescriptor(value, key)?.value;
}

function ownKeys(value: object): readonly PropertyKey[] | undefined {
  try {
    return Reflect.ownKeys(value);
  } catch {
    return undefined;
  }
}

function hasExactlyStringKeys(value: object, expected: readonly string[]): boolean {
  const keys = ownKeys(value);
  if (keys === undefined || keys.length !== expected.length) return false;
  const expectedSet = new Set(expected);
  for (const key of keys) {
    if (typeof key !== "string" || !expectedSet.has(key)) return false;
  }
  return true;
}

function boundedString(value: unknown, limit: number): value is string {
  if (typeof value !== "string" || value.length > limit) return false;
  let bytes = 0;
  for (let index = 0; index < value.length && bytes <= limit; index += 1) {
    const first = value.charCodeAt(index);
    if (first <= 0x7f) {
      bytes += 1;
    } else if (first <= 0x7ff) {
      bytes += 2;
    } else if (first >= 0xd800 && first <= 0xdbff) {
      const second = value.charCodeAt(index + 1);
      if (second >= 0xdc00 && second <= 0xdfff) {
        bytes += 4;
        index += 1;
      } else {
        bytes += 3;
      }
    } else {
      bytes += 3;
    }
  }
  return bytes <= limit;
}

function safeArrayLength(value: unknown): number | undefined {
  try {
    if (!Array.isArray(value)) return undefined;
    const descriptor = Object.getOwnPropertyDescriptor(value, "length");
    if (descriptor === undefined || !Object.prototype.hasOwnProperty.call(descriptor, "value")) {
      return undefined;
    }
    return typeof descriptor.value === "number" && Number.isSafeInteger(descriptor.value)
      ? descriptor.value
      : undefined;
  } catch {
    return undefined;
  }
}

/** Read an array slot without invoking an accessor or falling through a proxy get trap. */
function readOwnArrayElement(value: unknown, index: number): ArrayElementRead {
  if (!Array.isArray(value)) return { ok: false };
  const descriptor = ownDataDescriptor(value, String(index));
  return descriptor === undefined
    ? { ok: false }
    : { ok: true, value: descriptor.value };
}

function isPlanToolName(value: unknown): boolean {
  return value === WORKBOOK_PLAN_TOOL_NAME;
}

function inspectOpenAIPlanCalls(toolCalls: unknown):
  | { readonly kind: "none" }
  | { readonly kind: "rejected" }
  | { readonly kind: "plan"; readonly candidate: OpenAIPlanCandidate } {
  const length = safeArrayLength(toolCalls);
  if (length === undefined || length > MAX_PROTOCOL_TOOL_CALLS) return { kind: "rejected" };

  let planCandidate: OpenAIPlanCandidate | undefined;
  let planCount = 0;
  for (let index = 0; index < length; index += 1) {
    const callResult = readOwnArrayElement(toolCalls, index);
    if (!callResult.ok) return { kind: "rejected" };
    const call = callResult.value;
    if (!isPlainRecord(call)) return { kind: "rejected" };

    const functionValue = ownDataValue(call, "function");
    if (!isPlainRecord(functionValue)) return { kind: "rejected" };
    const name = ownDataValue(functionValue, "name");
    if (!isPlanToolName(name)) {
      const argumentsValue = ownDataValue(functionValue, "arguments");
      if (
        boundedString(argumentsValue, MAX_PROTOCOL_ARGUMENT_BYTES) &&
        hasPlanArgumentShape(argumentsValue)
      ) {
        return { kind: "rejected" };
      }
      continue;
    }

    planCount += 1;
    if (planCount > 1) return { kind: "rejected" };
    planCandidate = { payload: ownDataValue(functionValue, "arguments") };
  }

  if (planCandidate === undefined) return { kind: "none" };
  return { kind: "plan", candidate: planCandidate };
}

function validateOpenAIPlanEnvelope(
  toolCalls: unknown,
  candidate: OpenAIPlanCandidate,
): unknown | undefined {
  const length = safeArrayLength(toolCalls);
  if (length !== 1) return undefined;

  const callResult = readOwnArrayElement(toolCalls, 0);
  if (!callResult.ok) return undefined;
  const call = callResult.value;
  if (!isPlainRecord(call) || !hasExactlyStringKeys(call, ["id", "type", "function"])) {
    return undefined;
  }

  const id = ownDataValue(call, "id");
  const type = ownDataValue(call, "type");
  const functionValue = ownDataValue(call, "function");
  if (typeof id !== "string" || id.length === 0 || type !== "function" || !isPlainRecord(functionValue)) {
    return undefined;
  }
  if (!hasExactlyStringKeys(functionValue, ["name", "arguments"])) return undefined;
  if (ownDataValue(functionValue, "name") !== WORKBOOK_PLAN_TOOL_NAME) return undefined;

  const argumentsValue = candidate.payload;
  if (!boundedString(argumentsValue, MAX_PROTOCOL_ARGUMENT_BYTES)) return undefined;
  return parseDirectPayload(argumentsValue);
}

function parseDirectPayload(raw: string): unknown | undefined {
  let payload: unknown;
  try {
    payload = JSON.parse(raw);
  } catch {
    return undefined;
  }
  return isPlainRecord(payload) ? payload : undefined;
}

type ParsedToolCallTransport =
  | { readonly ok: true; readonly value: unknown }
  | { readonly ok: false };

/** Parse only the bounded primitive transport emitted by the direct agent. */
function parseToolCallTransport(value: unknown): ParsedToolCallTransport {
  if (typeof value !== "string" || !boundedString(value, MAX_PROTOCOL_TOOL_CALL_BYTES)) {
    return { ok: false };
  }
  try {
    return { ok: true, value: JSON.parse(value) };
  } catch {
    return { ok: false };
  }
}

interface XmlEnvelope {
  readonly payload: unknown;
  readonly hasXmlMarker: boolean;
  readonly malformed: boolean;
  readonly isPlanEnvelope: boolean;
}

function parseXmlFallback(content: string): XmlEnvelope {
  const markerOpen = content.indexOf("<tool_call>");
  const markerClose = content.indexOf("</tool_call>");
  if (markerOpen < 0 && markerClose < 0) {
    return { payload: undefined, hasXmlMarker: false, malformed: false, isPlanEnvelope: false };
  }

  const trimmed = content.trim();
  const open = "<tool_call>";
  const close = "</tool_call>";
  if (!trimmed.startsWith(open) || !trimmed.endsWith(close)) {
    return { payload: undefined, hasXmlMarker: true, malformed: true, isPlanEnvelope: false };
  }
  const body = trimmed.slice(open.length, trimmed.length - close.length).trim();
  if (
    body.length === 0 ||
    !boundedString(body, MAX_PROTOCOL_ARGUMENT_BYTES) ||
    body.includes("<tool_call") ||
    body.includes("</tool_call>") ||
    body.includes("<") ||
    body.includes(">")
  ) {
    return { payload: undefined, hasXmlMarker: true, malformed: true, isPlanEnvelope: false };
  }

  let envelope: unknown;
  try {
    envelope = JSON.parse(body);
  } catch {
    return { payload: undefined, hasXmlMarker: true, malformed: true, isPlanEnvelope: false };
  }
  if (!isPlainRecord(envelope) || !hasExactlyStringKeys(envelope, ["name", "arguments"])) {
    return { payload: undefined, hasXmlMarker: true, malformed: true, isPlanEnvelope: false };
  }
  if (ownDataValue(envelope, "name") !== WORKBOOK_PLAN_TOOL_NAME) {
    return { payload: undefined, hasXmlMarker: true, malformed: false, isPlanEnvelope: false };
  }
  const payload = ownDataValue(envelope, "arguments");
  if (!isPlainRecord(payload)) {
    return { payload: undefined, hasXmlMarker: true, malformed: true, isPlanEnvelope: true };
  }
  return { payload, hasXmlMarker: true, malformed: false, isPlanEnvelope: true };
}

function hasPlainJsonPlanShape(content: string): boolean {
  const trimmed = content.trim();
  if (trimmed.startsWith("{") || trimmed.startsWith("[")) return true;
  return (
    hasPlanBranchDiscriminators(content) &&
    (content.includes('"operations"') ||
      content.includes('"verification"') ||
      content.includes('"reads"') ||
      content.includes('"message"'))
  );
}

function hasPlanArgumentShape(content: string): boolean {
  return hasPlanBranchDiscriminators(content);
}

function hasPlanBranchDiscriminators(content: string): boolean {
  return (
    /"version"\s*:/.test(content) &&
    /"outcome"\s*:\s*"(?:execute|inspect|ask)"/.test(content)
  );
}

function deepFreeze<T>(value: T, seen = new Set<object>()): T {
  if (value === null || (typeof value !== "object" && typeof value !== "function")) return value;
  if (seen.has(value as object)) return value;
  seen.add(value as object);
  const keys = ownKeys(value as object);
  if (keys !== undefined) {
    for (const key of keys) {
      const descriptor = ownDataDescriptor(value as object, key);
      if (descriptor !== undefined) deepFreeze(descriptor.value, seen);
    }
  }
  return Object.freeze(value);
}

async function normalizePayload(
  payload: unknown,
  transport: "openai-tool-call" | "xml-fallback",
): Promise<WorkbookPlanProtocolResult> {
  if (!isPlainRecord(payload)) return reject();
  const normalized = await canonicalizeWorkbookPlan(payload);
  if (normalized.stage !== "success") return reject();

  const result: NormalizedWorkbookPlan = {
    ok: true,
    kind: "plan",
    transport,
    plan: deepFreeze(normalized.value),
    semanticDigest: normalized.semanticDigest,
  };
  return deepFreeze(result);
}

/**
 * Classify and normalize the sole V1 plan transport. A `none` result is
 * intentional: primitive compatibility responses remain on the existing
 * agent path and are not interpreted by this adapter.
 */
export async function normalizeWorkbookPlanResponse(
  content: unknown,
  serializedToolCalls: unknown,
): Promise<WorkbookPlanProtocolResult> {
  if (content !== null && !boundedString(content, MAX_PROTOCOL_CONTENT_BYTES)) return reject();
  const contentText = typeof content === "string" ? content : "";

  const parsedToolCalls = parseToolCallTransport(serializedToolCalls);
  if (!parsedToolCalls.ok) return reject();
  const toolCalls = parsedToolCalls.value;
  const openAI = inspectOpenAIPlanCalls(toolCalls);
  if (openAI.kind === "rejected") return reject();

  const xml = parseXmlFallback(contentText);
  if (xml.malformed) return reject();

  if (openAI.kind === "plan") {
    if (xml.hasXmlMarker || hasPlainJsonPlanShape(contentText) || contentText.trim() !== "") return reject();
    const payload = validateOpenAIPlanEnvelope(toolCalls, openAI.candidate);
    if (payload === undefined) return reject();
    return normalizePayload(payload, "openai-tool-call");
  }

  if (xml.hasXmlMarker) {
    if (!xml.isPlanEnvelope) {
      const toolCallLength = safeArrayLength(toolCalls);
      return toolCallLength === 0 ? none() : reject();
    }
    const toolCallLength = safeArrayLength(toolCalls);
    if (!xml.payload || toolCallLength !== 0) return reject();
    return normalizePayload(xml.payload, "xml-fallback");
  }

  if (hasPlainJsonPlanShape(contentText)) return reject();
  return none();
}

/**
 * Cheap, side-effect-free preclassification for the direct compatibility
 * adapter. It parses only the bounded serialized tool-call transport and
 * decides whether the strict plan protocol should be consulted instead of the
 * primitive path.
 */
export function isWorkbookPlanProtocolCandidate(
  content: unknown,
  serializedToolCalls: unknown,
): boolean {
  if (content !== null && typeof content !== "string") return true;
  const parsedToolCalls = parseToolCallTransport(serializedToolCalls);
  if (!parsedToolCalls.ok) return true;
  const toolCalls = parsedToolCalls.value;

  if (typeof content === "string") {
    if (!boundedString(content, MAX_PROTOCOL_CONTENT_BYTES)) return true;
    const xml = parseXmlFallback(content);
    if (xml.hasXmlMarker) {
      // A malformed XML marker is fail-closed. A complete envelope is
      // classified by its structural tool name, never by argument text.
      if (xml.malformed || xml.isPlanEnvelope) return true;
      const primitiveToolCallLength = safeArrayLength(toolCalls);
      return primitiveToolCallLength !== 0;
    }
    if (hasPlanArgumentShape(content)) return true;
  }

  const inspected = inspectOpenAIPlanCalls(toolCalls);
  return inspected.kind !== "none";
}
