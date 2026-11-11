const TRACE_KEY = "hermes.powerpoint.traces";
const MAX_TRACES = 50;
const MAX_METADATA_EVENTS = 100;
export interface ModelAttemptTrace { attempt: number; durationMs: number; statusCode?: number; category?: string; outcome: "success" | "error"; }
export interface ToolTrace { name: string; durationMs: number; ok: boolean; category?: string; }
export interface TurnTrace {
  id: string;
  startedAt: number;
  finishedAt?: number;
  effectiveIterationCap: number;
  modelCalls: number;
  toolCalls: number;
  errors: number;
  checkpoint?: "continue" | "stop";
  outcome?: string;
  modelAttempts: ModelAttemptTrace[];
  tools: ToolTrace[];
}
export interface ModelCallMetadata { attempt: number; durationMs: number; statusCode?: number; category?: string; outcome: "success" | "error"; }
export function startTrace(effectiveIterationCap: number): TurnTrace {
  return { id: `trace_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`, startedAt: Date.now(), effectiveIterationCap, modelCalls: 0, toolCalls: 0, errors: 0, modelAttempts: [], tools: [] };
}
export function recordModelCall(trace: TurnTrace, metadata: ModelCallMetadata | number, durationMs?: number, streamedOrStatus?: boolean | number, statusCode?: number): void {
  const event: ModelCallMetadata = typeof metadata === "number"
    ? { attempt: metadata + 1, durationMs: Math.max(0, Math.round(durationMs ?? 0)), statusCode: typeof streamedOrStatus === "number" ? streamedOrStatus : statusCode, outcome: "success" }
    : metadata;
  trace.modelCalls += 1;
  if (event.outcome === "error") trace.errors += 1;
  if (trace.modelAttempts.length < MAX_METADATA_EVENTS) trace.modelAttempts.push({ attempt: Math.max(1, Math.floor(event.attempt)), durationMs: Math.max(0, Math.round(event.durationMs)), ...(event.statusCode === undefined ? {} : { statusCode: event.statusCode }), ...(event.category ? { category: event.category.slice(0, 40) } : {}), outcome: event.outcome });
}
export function recordToolResult(trace: TurnTrace, name: string, durationMs: number, ok: boolean, category?: string): void {
  trace.toolCalls += 1;
  if (!ok) trace.errors += 1;
  if (trace.tools.length < MAX_METADATA_EVENTS) trace.tools.push({ name: name.slice(0, 80), durationMs: Math.max(0, Math.round(durationMs)), ok, ...(category ? { category: category.slice(0, 40) } : {}) });
}
export function recordCheckpoint(trace: TurnTrace, decision: "continue" | "stop"): void { trace.checkpoint = decision; }
export function finalizeTrace(trace: TurnTrace): void {
  trace.finishedAt = Date.now();
  try {
    const traces = loadTraces();
    traces.push(trace);
    localStorage.setItem(TRACE_KEY, JSON.stringify(traces.slice(-MAX_TRACES)));
  } catch {
    // Telemetry is best-effort and must never block task-pane cleanup.
  }
}
export function loadTraces(): TurnTrace[] { try { const value = JSON.parse(localStorage.getItem(TRACE_KEY) || "[]"); return Array.isArray(value) ? (value as TurnTrace[]).slice(-MAX_TRACES) : []; } catch { return []; } }
export function clearTraces(): void { localStorage.removeItem(TRACE_KEY); }
