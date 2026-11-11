/**
 * Metadata-only turn tracing for debugging without exposing sensitive content.
 *
 * NEVER records: workbook values, user prompts, tool arguments, provider
 * response bodies, API keys, or any content that could contain PII or
 * proprietary workbook data.
 *
 * Only records: correlation IDs, timing, result categories, and counts.
 * Retention is bounded to MAX_TRACES (FIFO eviction in localStorage).
 */

import { ToolErrorCategory } from "./toolErrors";

const STORAGE_KEY = "hermes.excel.traces";
const MAX_TRACES = 50;

export interface TraceToolResult {
  /** Tool name (e.g. "read_range", "write_range"). */
  name: string;
  /** Execution duration in ms. */
  durationMs: number;
  /** True when ok:false or an exception was caught. */
  isError: boolean;
  /** Error category if isError, undefined otherwise. */
  errorCategory?: ToolErrorCategory;
}

export interface TraceModelCall {
  /** 0-based attempt index (0 = first try). */
  attempt: number;
  /** Duration in ms. */
  durationMs: number;
  /** HTTP status code, if the provider returned one. */
  statusCode?: number;
  /** True if the call succeeded. */
  success: boolean;
}

export interface TurnTrace {
  /** Unique correlation ID for this turn (timestamp-based). */
  turnId: string;
  /** ISO-8601 timestamp of turn start. */
  startedAt: string;
  /** Total turn duration in ms. */
  totalDurationMs: number;
  /** Number of model-call attempts (including retries). */
  modelCallAttempts: number;
  /** Per-call metadata. */
  modelCalls: TraceModelCall[];
  /** Number of tool executions attempted. */
  toolCallCount: number;
  /** Per-tool result metadata. */
  toolResults: TraceToolResult[];
  /** How the turn ended. */
  outcome: "completed" | "stopped_user" | "stopped_checkpoint" | "iteration_limit" | "timeout" | "error";
  /** Error category if outcome is "error". */
  errorCategory?: ToolErrorCategory;
  /** Error message summary (no content, just type). */
  errorSummary?: string;
  /** Whether the 30s checkpoint fired. */
  checkpointFired: boolean;
  /** Effective max iterations used. */
  maxIterations: number;
  /** Turn number (incremented per session). */
  turnNumber: number;
}

let turnCounter = loadTurnCounter();

function loadTurnCounter(): number {
  try {
    const raw = localStorage.getItem("hermes.excel.turnCounter");
    return raw ? Math.max(0, parseInt(raw, 10) || 0) : 0;
  } catch {
    return 0;
  }
}

function saveTurnCounter(value: number): void {
  try {
    localStorage.setItem("hermes.excel.turnCounter", String(value));
  } catch {
    // localStorage may be unavailable in some contexts
  }
}

/**
 * Start a new trace. Call at the beginning of a turn.
 * Returns a mutable trace object that the agent loop fills in.
 */
export function startTrace(maxIterations: number): TurnTrace {
  turnCounter += 1;
  saveTurnCounter(turnCounter);
  return {
    turnId: `turn_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    startedAt: new Date().toISOString(),
    totalDurationMs: 0,
    modelCallAttempts: 0,
    modelCalls: [],
    toolCallCount: 0,
    toolResults: [],
    outcome: "completed",
    checkpointFired: false,
    maxIterations,
    turnNumber: turnCounter,
  };
}

/**
 * Record a model call attempt. Call after each fetch (success or failure).
 */
export function recordModelCall(
  trace: TurnTrace,
  attempt: number,
  durationMs: number,
  success: boolean,
  statusCode?: number,
): void {
  trace.modelCallAttempts += 1;
  trace.modelCalls.push({ attempt, durationMs, statusCode, success });
}

/**
 * Record a tool execution result.
 */
export function recordToolResult(
  trace: TurnTrace,
  name: string,
  durationMs: number,
  isError: boolean,
  errorCategory?: ToolErrorCategory,
): void {
  trace.toolCallCount += 1;
  trace.toolResults.push({ name, durationMs, isError, errorCategory });
}

/**
 * Finalize and persist a trace. Call at the end of a turn.
 */
export function finalizeTrace(trace: TurnTrace): void {
  trace.totalDurationMs = Date.now() - new Date(trace.startedAt).getTime();
  persistTrace(trace);
}

function persistTrace(trace: TurnTrace): void {
  try {
    const existing = loadTraces();
    existing.push(trace);
    // FIFO eviction
    while (existing.length > MAX_TRACES) {
      existing.shift();
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(existing));
  } catch {
    // Fail silently — tracing is best-effort
  }
}

/**
 * Load all persisted traces (newest last).
 */
export function loadTraces(): TurnTrace[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed as TurnTrace[];
  } catch {
    return [];
  }
}

/**
 * Clear all traces. Call this when the user wants to purge trace data.
 */
export function clearTraces(): void {
  try {
    localStorage.removeItem(STORAGE_KEY);
  } catch {
    // Best-effort
  }
}
