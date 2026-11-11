import { executePowerPointToolParsed, getPresentationContext, toolSchemas } from "./powerPointTools";
import { validateToolArguments } from "./toolValidation";
import { asToolError, internalError, officeHostError, protocolError, providerError, repetitionError, stringifyToolError, validationError, type ToolError } from "./toolErrors";
import { HermesSettings } from "./settings";
import { ActivityItem, ChatMessage, ToolCall } from "./types";

export const MODEL_CALL_TIMEOUT_MS = 20_000;
export const TURN_TIMEOUT_MS = 90_000;
/** A native request runs the entire server agent, not one model completion. */
export const NATIVE_TURN_TIMEOUT_MS = 300_000;
export const CHECKPOINT_TIMEOUT_MS = 30_000;
export const MAX_RETRIES = 2;
export const ABSOLUTE_MAX_ITERATIONS = 20;
export interface ModelCallEvent { attempt: number; durationMs: number; statusCode?: number; category?: string; outcome: "success" | "error"; }
export interface ToolResultEvent { name: string; durationMs: number; ok: boolean; category?: string; }
export interface AgentCallbacks {
  onAssistantStart: () => void;
  onAssistantDelta: (delta: string) => void;
  onAssistantFinal: (content: string) => void;
  onActivity: (item: ActivityItem) => void;
  onStatus: (status: string) => void;
  onCheckpoint?: (signal: AbortSignal) => Promise<"continue" | "stop">;
  onCheckpointResult?: (decision: "continue" | "stop") => void;
  onModelCall?: (event: ModelCallEvent) => void;
  onToolResult?: (event: ToolResultEvent) => void;
  onResponseHeaders?: (headers: Headers, status: number) => void;
}
export interface AgentRunResult { messages: ChatMessage[]; finalContent: string; stopReason?: "user" | "timeout" | "iteration_limit" | "checkpoint"; }
export interface AgentRunOptions { bridgeSessionId?: string | null; }
interface StreamResult { content: string; toolCalls: ToolCall[]; emitted: boolean; statusCode?: number; }
const OFFICE_OPERATION_IN_FLIGHT_MESSAGE = "A PowerPoint operation may still complete. Verify the deck state before starting another request.";
const EXTERNAL_OPERATION_IN_FLIGHT_MESSAGE = "A previous Hermes operation is still settling. Wait before starting another request.";
export type PendingOperationKind = "context" | "checkpoint" | "model" | "read" | "office" | "cleanup";
export type ExternalOperationTracker = <T>(operation: Promise<T>, kind: PendingOperationKind) => Promise<T>;

const SYSTEM_PROMPT = `You are Hermes, an agentic AI assistant running inside the user's active Microsoft PowerPoint presentation.
Inspect before editing, make small verifiable changes, and use tool errors to correct your next action. Use the current tool schemas exactly. If native tool calls are unavailable, emit <tool_call>{"name":"tool_name","arguments":{}}</tool_call>.`;

export class HermesAgent {
  private messages: ChatMessage[] = [];
  private activeGeneration: string | null = null;
  private pendingOperations = new Map<Promise<unknown>, PendingOperationKind>();
  get history(): ChatMessage[] { return [...this.messages]; }
  reset(): void { this.activeGeneration = null; this.messages = []; }
  private trackOperation<T>(operation: Promise<T>, kind: PendingOperationKind): Promise<T> {
    this.pendingOperations.set(operation, kind);
    operation.then(
      () => { this.pendingOperations.delete(operation); },
      () => { this.pendingOperations.delete(operation); },
    );
    return operation;
  }
  private pendingOperationError(): Error | null {
    if (!this.pendingOperations.size) return null;
    if ([...this.pendingOperations.values()].includes("office")) {
      const toolError = officeHostError(OFFICE_OPERATION_IN_FLIGHT_MESSAGE);
      return Object.assign(new Error(toolError.error), { toolError });
    }
    const toolError = internalError(EXTERNAL_OPERATION_IN_FLIGHT_MESSAGE);
    return Object.assign(new Error(toolError.error), { toolError });
  }
  async run(userRequest: string, settings: HermesSettings, callbacks: AgentCallbacks, userSignal: AbortSignal, options: AgentRunOptions = {}): Promise<AgentRunResult> {
    if (this.activeGeneration) throw new Error("A PowerPoint Hermes turn is already running.");
    const pendingError = this.pendingOperationError();
    if (pendingError) throw pendingError;
    const generation = makeCorrelationId(); this.activeGeneration = generation;
    const nativeMode = isNativeHermesMode(settings);
    const bridgeSessionId = options.bridgeSessionId?.trim() || null;
    if (nativeMode && (!bridgeSessionId || !isUuid(bridgeSessionId))) {
      const toolError = validationError("A live PowerPoint bridge session is required for native Hermes mode.");
      this.activeGeneration = null;
      throw Object.assign(new Error(toolError.error), { toolError });
    }
    const turnStarted = Date.now();
    const turnBudgetMs = nativeMode ? NATIVE_TURN_TIMEOUT_MS : TURN_TIMEOUT_MS;
    let turnDeadline = turnStarted + turnBudgetMs;
    const turnController = new AbortController();
    const timeoutTurn = () => turnController.abort(new DOMException("The turn exceeded its time limit.", "TimeoutError"));
    let turnTimer = setTimeout(timeoutTurn, turnBudgetMs);
    const extendTurn = () => {
      turnDeadline = Date.now() + turnBudgetMs;
      clearTimeout(turnTimer);
      turnTimer = setTimeout(timeoutTurn, turnBudgetMs);
    };
    const checkpointController = new AbortController();
    const unlinkCheckpoint = linkAbortSignals(turnController.signal, checkpointController);
    let checkpointStopped = false;
    let checkpointFailure: unknown;
    let nativeCheckpoint: Promise<void> | undefined;
    const unlink = linkAbortSignals(userSignal, turnController); const startedHistory = [...this.messages];
    let presentationContext = ""; let checkpointFired = false;
    const current = () => this.activeGeneration === generation;
    const ensureCurrent = () => { if (!current()) throw new DOMException("The running Hermes generation is stale.", "AbortError"); };
    const notify = <T extends unknown[]>(callback: ((...args: T) => void) | undefined, ...args: T): void => { if (current() && callback) callback(...args); };
    const maybeCheckpoint = async (): Promise<"continue" | "stop"> => {
      if (checkpointFired || !callbacks.onCheckpoint || Date.now() - turnStarted < CHECKPOINT_TIMEOUT_MS) return "continue";
      checkpointFired = true;
      const checkpointOperation = this.trackOperation(Promise.resolve().then(() => callbacks.onCheckpoint!(checkpointController.signal)), "checkpoint");
      const decision = await raceWithAbort(checkpointOperation, turnController.signal, userSignal);
      ensureCurrent();
      if (decision === "continue") extendTurn();
      safeNotify(callbacks.onCheckpointResult, decision); return decision;
    };
    const stopAtCheckpoint = (decision: "continue" | "stop"): AgentRunResult | null => {
      if (decision !== "stop") return null;
      const finalContent = "Turn stopped at the 30-second checkpoint.";
      notify(callbacks.onAssistantFinal, finalContent); notify(callbacks.onStatus, "Stopped at checkpoint");
      return { messages: this.history, finalContent, stopReason: "checkpoint" };
    };
    // A native model call stays in flight for the whole turn: checkpoints and
    // progress must run concurrently, not only between model calls.
    const checkpointTimer = nativeMode ? setTimeout(() => {
      if (!current() || turnController.signal.aborted) return;
      nativeCheckpoint = maybeCheckpoint().then((decision) => {
        if (decision === "stop") {
          checkpointStopped = true;
          turnController.abort(new DOMException("Stopped at checkpoint.", "AbortError"));
        }
      }).catch((error) => {
        checkpointFailure = error;
        turnController.abort(new DOMException("Checkpoint failed.", "AbortError"));
      });
    }, CHECKPOINT_TIMEOUT_MS) : undefined;
    const progressTimer = nativeMode ? setInterval(() => {
      if (current()) safeNotify(callbacks.onStatus, `Hermes is working (${Math.floor((Date.now() - turnStarted) / 1000)}s)…`);
    }, 5_000) : undefined;
    try {
      if (userSignal.aborted) throw abortReason(userSignal);
      const contextOperation = this.trackOperation(Promise.resolve().then(() => getPresentationContext()), "context");
      presentationContext = await raceWithAbort(contextOperation, turnController.signal, userSignal);
      ensureCurrent(); throwIfAborted(turnController.signal, userSignal);
      const workingMessages = this.buildTurnMessages(userRequest, presentationContext, startedHistory);
      const toolCallCounts = new Map<string, number>();
      const maxIterations = Math.min(Math.max(Math.floor(Number(settings.maxIterations) || 1), 1), ABSOLUTE_MAX_ITERATIONS);
      notify(callbacks.onAssistantStart);
      for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
        ensureCurrent(); throwIfAborted(turnController.signal, userSignal); notify(callbacks.onStatus, `Thinking ${iteration}/${maxIterations}`);
        const beforeModel = stopAtCheckpoint(await maybeCheckpoint()); if (beforeModel) return beforeModel;
        const response = await this.chatCompletion(settings, workingMessages, turnController.signal, userSignal, turnDeadline, callbacks, nativeMode, bridgeSessionId);
        if (nativeCheckpoint) await nativeCheckpoint;
        ensureCurrent(); throwIfAborted(turnController.signal, userSignal);
        const afterModel = stopAtCheckpoint(await maybeCheckpoint()); if (afterModel) return afterModel;
        const toolCalls = response.toolCalls.length ? response.toolCalls : parseXmlToolCalls(response.content, generation);
        if (!toolCalls.length) {
          const finalContent = removeXmlToolCalls(response.content).trim(); notify(callbacks.onAssistantFinal, finalContent); notify(callbacks.onStatus, "Ready");
          this.commitIfCurrent(generation, userRequest, presentationContext, finalContent, startedHistory); return { messages: this.history, finalContent };
        }
        if (nativeMode) {
          const finalContent = removeXmlToolCalls(response.content).trim();
          notify(callbacks.onAssistantFinal, finalContent); notify(callbacks.onStatus, "Ready");
          this.commitIfCurrent(generation, userRequest, presentationContext, finalContent, startedHistory);
          return { messages: this.history, finalContent };
        }
        workingMessages.push({ role: "assistant", content: removeXmlToolCalls(response.content).trim() || null, tool_calls: toolCalls });
        for (const toolCall of toolCalls) {
          ensureCurrent(); throwIfAborted(turnController.signal, userSignal);
          const schema = toolSchemas.find((entry) => entry.function.name === toolCall.function.name)?.function.parameters;
          const parsedValidation = validateToolArguments(toolCall.function.arguments, schema as Record<string, unknown> | undefined, toolCall.function.name);
          const signature = parsedValidation.ok ? `${toolCall.function.name}:${canonicalizeJson(parsedValidation.parsedArgs)}` : `${toolCall.function.name}:invalid:${canonicalizeJson(toolCall.function.arguments)}`;
          const count = (toolCallCounts.get(signature) ?? 0) + 1; toolCallCounts.set(signature, count); const label = `${toolCall.function.name} (${count})`;
          notify(callbacks.onActivity, { id: toolCall.id, status: "running", label }); const toolStarted = Date.now(); let result = ""; let toolOk = false; let toolCategory: string | undefined;
          try {
            if (count > 2) { result = stringifyToolError(repetitionError("Repeated identical tool call blocked; use the prior result.")); toolCategory = "repetition"; }
            else if (!parsedValidation.ok) { result = stringifyToolError(parsedValidation.error); toolCategory = parsedValidation.error.category; }
            else {
              const officePromise = this.trackOperation(Promise.resolve().then(() => executePowerPointToolParsed(toolCall.function.name, parsedValidation.parsedArgs)), "office");
              result = await raceWithAbort(officePromise, turnController.signal, userSignal);
              const parsedResult = safeParse(result) as { ok?: unknown; category?: unknown } | null;
              toolOk = parsedResult?.ok !== false;
              toolCategory = typeof parsedResult?.category === "string" ? parsedResult.category : undefined;
            }
          } catch (error) {
            toolCategory = classifyTurnError(error, userSignal, turnController.signal).category;
            throw error;
          } finally { safeNotify(callbacks.onToolResult, { name: toolCall.function.name, durationMs: Date.now() - toolStarted, ok: toolOk, category: toolCategory }); }
          const parsedResult = safeParse(result) as { ok?: unknown } | null; notify(callbacks.onActivity, { id: toolCall.id, status: parsedResult?.ok === false ? "error" : "done", label }); workingMessages.push({ role: "tool", tool_call_id: toolCall.id, name: toolCall.function.name, content: result });
        }
      }
      const finalContent = `I reached the maximum of ${maxIterations} agent iterations and stopped to avoid unbounded presentation changes.`; notify(callbacks.onAssistantFinal, finalContent); notify(callbacks.onStatus, "Stopped at iteration limit"); this.commitIfCurrent(generation, userRequest, presentationContext, finalContent, startedHistory); return { messages: this.history, finalContent, stopReason: "iteration_limit" };
    } catch (error) {
      if (checkpointStopped && !userSignal.aborted) return stopAtCheckpoint("stop")!;
      const toolError = classifyTurnError(checkpointFailure ?? error, userSignal, checkpointFailure ? new AbortController().signal : turnController.signal);
      if (toolError.category === "cancellation") throw Object.assign(new DOMException(toolError.error, "AbortError"), { toolError });
      if (toolError.category === "timeout") throw Object.assign(new DOMException(toolError.error, "TimeoutError"), { toolError });
      const safeError = new Error(toolError.error);
      if (toolError.name) safeError.name = toolError.name;
      if (toolError.statusCode !== undefined) Object.assign(safeError, { statusCode: toolError.statusCode });
      if (error && typeof error === "object" && (error as Record<string, unknown>).partial === true) Object.assign(safeError, { partial: true });
      throw Object.assign(safeError, { toolError });
    } finally {
      clearTimeout(turnTimer); clearTimeout(checkpointTimer); clearInterval(progressTimer);
      checkpointController.abort(); unlinkCheckpoint(); unlink();
      if (this.activeGeneration === generation) this.activeGeneration = null;
    }
  }
  private commitIfCurrent(generation: string, request: string, context: string, content: string, base: ChatMessage[]): void { if (this.activeGeneration === generation) this.messages = [...base, this.userHistoryMessage(request, context), { role: "assistant", content }]; }
  private buildTurnMessages(request: string, context: string, base: ChatMessage[]): ChatMessage[] { return [{ role: "system", content: SYSTEM_PROMPT }, ...base.slice(-12), this.userHistoryMessage(request, context)]; }
  private userHistoryMessage(request: string, context: string): ChatMessage { return { role: "user", content: `Fresh presentation context:\n${context}\n\nUser request:\n${request}` }; }

  private async chatCompletion(settings: HermesSettings, messages: ChatMessage[], turnSignal: AbortSignal, userSignal: AbortSignal, turnDeadline: number, callbacks: AgentCallbacks, nativeMode: boolean, bridgeSessionId: string | null): Promise<StreamResult> {
    const url = chatCompletionsUrl(settings.baseUrl);
    const headers: Record<string, string> = { "Content-Type": "application/json", ...(settings.requestHeaders ?? {}) };
    if (nativeMode && bridgeSessionId) headers["X-Hermes-Session-Id"] = bridgeSessionId;
    const body: Record<string, unknown> = { model: settings.model, messages, temperature: 0.2 };
    if (!nativeMode) { body.tools = toolSchemas; body.tool_choice = "auto"; }
    const remainingTurn = turnDeadline - Date.now();
    if (remainingTurn <= 0) throw new DOMException("The turn exceeded its time limit.", "TimeoutError");
    const callBudgetMs = nativeMode ? remainingTurn : Math.min(MODEL_CALL_TIMEOUT_MS, remainingTurn);
    const callDeadline = Date.now() + callBudgetMs;
    // Native calls use the extendable turn signal; a second fixed timer would
    // incorrectly kill a turn after the user chose Continue.
    const call = nativeMode ? { controller: new AbortController(), cleanup: () => {} } : deadlineSignal(callBudgetMs);
    const unlink = linkAbortSignals(turnSignal, call.controller);
    let lastError: unknown;
    try {
      for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
        throwIfAborted(call.controller.signal, userSignal);
        const started = Date.now();
        try {
          const result = await this.oneModelCall(url, headers, body, call.controller.signal, userSignal, callbacks, nativeMode);
          safeNotify(callbacks.onModelCall, { attempt: attempt + 1, durationMs: Date.now() - started, statusCode: result.statusCode ?? 200, outcome: "success" });
          return result;
        } catch (error) {
          lastError = error;
          const classified = classifyTurnError(error, userSignal, call.controller.signal);
          const statusCode = statusCodeOf(error) ?? classified.statusCode;
          safeNotify(callbacks.onModelCall, { attempt: attempt + 1, durationMs: Date.now() - started, statusCode, category: classified.category, outcome: "error" });
          const partial = error && typeof error === "object" && (error as Record<string, unknown>).partial === true;
          const retryable = classified.category === "network" || statusCode === 429 || (statusCode !== undefined && statusCode >= 500 && statusCode <= 599);
          const delay = 250 * 2 ** attempt;
          const remaining = callDeadline - Date.now();
          if (nativeMode || !retryable || partial || attempt >= MAX_RETRIES || remaining <= delay) {
            if (retryable && remaining <= delay && !userSignal.aborted && !call.controller.signal.aborted) throw new DOMException("The model call has insufficient time for another retry.", "TimeoutError");
            throw error;
          }
          await interruptibleSleep(delay, call.controller.signal, userSignal);
        }
      }
      throw lastError;
    } finally {
      unlink();
      call.cleanup();
    }
  }
  private async oneModelCall(url: string, headers: Record<string, string>, baseBody: Record<string, unknown>, callSignal: AbortSignal, userSignal: AbortSignal, callbacks: AgentCallbacks, nativeMode: boolean): Promise<StreamResult> {
    let response: Response;
    try { response = await raceWithAbort(this.trackOperation(Promise.resolve().then(() => fetch(url, { method: "POST", headers, signal: callSignal, body: JSON.stringify({ ...baseBody, stream: true }) })), "model"), callSignal, userSignal); }
    catch (error) { throw classifyFetchError(error, userSignal, callSignal); }
    this.observeResponse(response, callbacks, callSignal, userSignal);
    if (!response.ok) { const text = await readResponseBody(response, callSignal, 2_000_000, this.trackOperation.bind(this)); const parsed = tryParseNonStreamResponse(text); if (parsed) { if (parsed.content) callbacks.onAssistantDelta(parsed.content); return { ...parsed, statusCode: response.status, emitted: Boolean(parsed.content || parsed.toolCalls.length) }; } throw providerError(`Model endpoint returned ${response.status}.`, response.status); }
    if (isJsonContentType(response.headers.get("content-type"))) { const text = await readResponseBody(response, callSignal, 2_000_000, this.trackOperation.bind(this)); const parsed = tryParseNonStreamResponse(text); if (!parsed) throw protocolError("Model endpoint returned invalid JSON."); if (parsed.content) callbacks.onAssistantDelta(parsed.content); return { ...parsed, statusCode: response.status, emitted: Boolean(parsed.content || parsed.toolCalls.length) }; }
    const emptyResponse = (): Promise<StreamResult> => {
      // A native request may already have edited the deck. Even a zero-byte
      // response is not evidence that dispatch failed: never replay it as JSON.
      if (nativeMode) throw protocolError("Hermes returned no completion. Presentation changes may have occurred; review the deck before retrying.");
      return this.fallbackNonStream(url, headers, baseBody, callSignal, userSignal, callbacks);
    };
    if (!response.body) return emptyResponse();
    const stream = await readOpenAIStream(response.body, callbacks, callSignal, this.trackOperation.bind(this));
    if (!stream.content && stream.toolCalls.length === 0) {
      if (stream.emitted) throw Object.assign(new Error("The model stream ended after partial output."), { partial: true, category: "protocol" });
      return emptyResponse();
    }
    return { ...stream, statusCode: response.status };
  }
  private async fallbackNonStream(url: string, headers: Record<string, string>, baseBody: Record<string, unknown>, signal: AbortSignal, userSignal: AbortSignal, callbacks: AgentCallbacks): Promise<StreamResult> {
    let response: Response; try { response = await raceWithAbort(this.trackOperation(Promise.resolve().then(() => fetch(url, { method: "POST", headers, signal, body: JSON.stringify({ ...baseBody, stream: false }) })), "model"), signal, userSignal); } catch (error) { throw classifyFetchError(error, userSignal, signal); }
    this.observeResponse(response, callbacks, signal, userSignal);
    const text = await readResponseBody(response, signal, 2_000_000, this.trackOperation.bind(this)); const parsed = tryParseNonStreamResponse(text); if (!parsed) throw Object.assign(protocolError(`Model endpoint returned invalid non-stream response (${response.status}).`), { statusCode: response.status }); if (parsed.content) callbacks.onAssistantDelta(parsed.content); return { ...parsed, statusCode: response.status, emitted: Boolean(parsed.content || parsed.toolCalls.length) };
  }
  private observeResponse(response: Response, callbacks: AgentCallbacks, signal: AbortSignal, userSignal: AbortSignal): void {
    try {
      callbacks.onResponseHeaders?.(response.headers, response.status);
      throwIfAborted(signal, userSignal);
    } catch (error) {
      // Metadata can cancel a turn before a reader exists; release that body too.
      if (response.body) {
        const body = response.body;
        void this.trackOperation(Promise.resolve().then(() => body.cancel()), "cleanup").catch(() => {});
      }
      throw error;
    }
  }
}

export function canonicalizeJson(value: unknown): string { const canonical = (input: unknown): string => { if (Array.isArray(input)) return `[${input.map(canonical).join(",")}]`; if (input && typeof input === "object") return `{${Object.keys(input as Record<string, unknown>).sort().map((key) => `${JSON.stringify(key)}:${canonical((input as Record<string, unknown>)[key])}`).join(",")}}`; return JSON.stringify(input); }; if (typeof value === "string") { const trimmed = value.trim(); try { return canonical(JSON.parse(trimmed)); } catch { return `raw:${trimmed}`; } } return canonical(value); }
export function generateXmlCallId(runId: string, sequence: number): string { return `xml_call_${runId}_${sequence}`; }
export function makeToolCallSignature(toolCall: ToolCall): string { return `${toolCall.function.name}:${canonicalizeJson(toolCall.function.arguments)}`; }
function tryParseNonStreamResponse(bodyText: string): StreamResult | null { const parsed = safeParse(bodyText) as Record<string, unknown> | null; const choices = parsed?.choices; if (!Array.isArray(choices) || !choices.length) return null; const message = (choices[0] as Record<string, unknown>)?.message as Record<string, unknown> | undefined; if (!message) return null; const toolCalls: ToolCall[] = []; for (const raw of (message.tool_calls as unknown[] | undefined) ?? []) { const item = raw as Record<string, unknown>; const fn = item.function as Record<string, unknown> | undefined; if (item.type === "function" && typeof fn?.name === "string") toolCalls.push({ id: typeof item.id === "string" ? item.id : `call_${toolCalls.length}`, type: "function", function: { name: fn.name, arguments: typeof fn.arguments === "string" ? fn.arguments : JSON.stringify(fn.arguments ?? {}) } }); } return { content: typeof message.content === "string" ? message.content : "", toolCalls, emitted: false }; }
function safeParse(value: string): unknown | null { try { return JSON.parse(value); } catch { return null; } }
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
function isUuid(value: string): boolean { return UUID_PATTERN.test(value); }
function isJsonContentType(contentType: string | null): boolean { return Boolean(contentType && /(?:^|;)\s*application\/json\s*(?:;|$)/i.test(contentType)); }
export async function raceWithAbort<T>(promise: Promise<T>, signal: AbortSignal, userSignal: AbortSignal): Promise<T> { if (userSignal.aborted) throw abortReason(userSignal); if (signal.aborted) throw abortReason(signal); return new Promise<T>((resolve, reject) => { const cleanup = () => { signal.removeEventListener("abort", onAbort); userSignal.removeEventListener("abort", onUserAbort); }; const onAbort = () => { cleanup(); reject(abortReason(signal)); }; const onUserAbort = () => { cleanup(); reject(abortReason(userSignal)); }; signal.addEventListener("abort", onAbort, { once: true }); userSignal.addEventListener("abort", onUserAbort, { once: true }); Promise.resolve(promise).then((value) => { cleanup(); resolve(value); }, (error) => { cleanup(); reject(error); }); }); }
function throwIfAborted(signal: AbortSignal, userSignal?: AbortSignal): void { if (userSignal?.aborted) throw abortReason(userSignal); if (signal.aborted) throw abortReason(signal); }
function abortReason(signal: AbortSignal): DOMException { const reason = signal.reason; return reason instanceof DOMException ? reason : new DOMException("The running Hermes task was stopped.", "AbortError"); }
function classifyFetchError(error: unknown, userSignal: AbortSignal, signal: AbortSignal): ToolError { if (userSignal.aborted) return asToolError(new DOMException("The running Hermes task was stopped.", "AbortError")); if (signal.aborted && signal.reason?.name === "TimeoutError") return asToolError(new DOMException("The operation exceeded its time limit.", "TimeoutError")); return error instanceof TypeError ? asToolError(error, "network", { fetchTypeError: true }) : asToolError(error, "internal", { fetchTypeError: false }); }
function classifyTurnError(error: unknown, userSignal: AbortSignal, signal: AbortSignal): ToolError {
  if (userSignal.aborted) return asToolError(new DOMException("The running Hermes task was stopped.", "AbortError"));
  if (signal.aborted && signal.reason?.name === "TimeoutError") return asToolError(new DOMException("The operation exceeded its time limit.", "TimeoutError"));
  if (error && typeof error === "object" && "toolError" in error) return asToolError((error as { toolError: unknown }).toolError);
  return asToolError(error, "internal", { fetchTypeError: false });
}
function statusCodeOf(error: unknown): number | undefined { return error && typeof error === "object" && typeof (error as Record<string, unknown>).statusCode === "number" ? (error as Record<string, unknown>).statusCode as number : undefined; }
function makeCorrelationId(): string { return `run_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`; }
function safeNotify<T extends unknown[]>(callback: ((...args: T) => void) | undefined, ...args: T): void { try { callback?.(...args); } catch { /* metadata must never change orchestration */ } }
export function isBlockedRepeatedToolCall(count: number): boolean { return count > 2; }
export function chatCompletionsUrl(baseUrl: string): string { const trimmed = baseUrl.trim().replace(/\/+$/, ""); const without = trimmed.replace(/\/chat\/completions$/i, ""); if (without === "http://127.0.0.1:8643/v1" || without === "http://localhost:8643/v1") return "/hermes/v1/chat/completions"; return trimmed.endsWith("/chat/completions") ? trimmed : `${trimmed}/chat/completions`; }
export function isNativeHermesMode(settings: Pick<HermesSettings, "baseUrl" | "model">): boolean {
  if (settings.model.trim() !== "hermes-agent") return false;
  const base = settings.baseUrl.trim().replace(/\/+$/, "");
  if (base === "/hermes/v1") return true;
  try {
    const parsed = new URL(base);
    const loopback = parsed.hostname === "127.0.0.1" || parsed.hostname === "localhost" || parsed.hostname === "::1";
    const deskPort = ["47650", "47651", "47652", "47653", "47654"].includes(parsed.port);
    const verifiedDesk = parsed.protocol === "https:" && parsed.hostname === "localhost" && deskPort;
    const legacyLoopback = parsed.protocol === "http:" && loopback && (parsed.port === "8646" || parsed.port === "8643" || deskPort);
    return (verifiedDesk || legacyLoopback) && parsed.pathname === "/v1" && !parsed.username && !parsed.password && !parsed.search && !parsed.hash;
  } catch {
    return false;
  }
}
export function deadlineSignal(ms: number): { controller: AbortController; cleanup: () => void } { const controller = new AbortController(); const timer = setTimeout(() => controller.abort(new DOMException("The model call exceeded its deadline.", "TimeoutError")), ms); return { controller, cleanup: () => clearTimeout(timer) }; }
export function linkAbortSignals(primary: AbortSignal, secondary: AbortController): () => void { const abort = () => secondary.abort(primary.reason ?? new DOMException("The running Hermes task was stopped.", "AbortError")); if (primary.aborted) abort(); else primary.addEventListener("abort", abort, { once: true }); return () => primary.removeEventListener("abort", abort); }
export async function interruptibleSleep(ms: number, signal: AbortSignal, userSignal?: AbortSignal): Promise<void> { await new Promise<void>((resolve, reject) => { let settled = false; let timer: ReturnType<typeof setTimeout>; const abort = () => finish(abortReason(userSignal?.aborted ? userSignal : signal)); const cleanup = () => { clearTimeout(timer); signal.removeEventListener("abort", abort); userSignal?.removeEventListener("abort", abort); }; const finish = (error?: unknown) => { if (settled) return; settled = true; cleanup(); error === undefined ? resolve() : reject(error); }; timer = setTimeout(() => finish(), ms); if (signal.aborted || userSignal?.aborted) abort(); else { signal.addEventListener("abort", abort, { once: true }); userSignal?.addEventListener("abort", abort, { once: true }); } }); }
function bestEffortCancel<T>(reader: ReadableStreamDefaultReader<T>, trackOperation?: ExternalOperationTracker): () => void {
  let cancellation: Promise<void> | null = null;
  return () => {
    if (!cancellation) {
      const cancelOperation = Promise.resolve().then(() => reader.cancel());
      cancellation = trackOperation ? trackOperation(cancelOperation, "cleanup") : cancelOperation;
      cancellation.then(
        () => undefined,
        () => undefined,
      );
    }
  };
}

function raceReaderRead<T>(
  reader: ReadableStreamDefaultReader<T>,
  signal: AbortSignal,
  cancel: () => void,
  trackOperation?: ExternalOperationTracker,
): Promise<ReadableStreamReadResult<T>> {
  if (signal.aborted) {
    cancel();
    return Promise.reject(abortReason(signal));
  }
  return new Promise<ReadableStreamReadResult<T>>((resolve, reject) => {
    let settled = false;
    const cleanup = () => signal.removeEventListener("abort", onAbort);
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      cleanup();
      callback();
    };
    const onAbort = () => {
      cancel();
      finish(() => reject(abortReason(signal)));
    };
    signal.addEventListener("abort", onAbort, { once: true });
    const originalRead = Promise.resolve().then(() => reader.read());
    const readPromise = trackOperation ? trackOperation(originalRead, "read") : originalRead;
    readPromise.then(
      (value) => finish(() => resolve(value)),
      (error) => finish(() => reject(error)),
    );
    if (signal.aborted) onAbort();
  });
}

export async function readResponseBody(response: Response, signal: AbortSignal, maxBytesOrTracker: number | ExternalOperationTracker = 2_000_000, trackOperation?: ExternalOperationTracker): Promise<string> {
  const maxBytes = typeof maxBytesOrTracker === "number" ? maxBytesOrTracker : 2_000_000;
  const tracker = typeof maxBytesOrTracker === "function" ? maxBytesOrTracker : trackOperation;
  if (!response.body) return "";
  const reader = response.body.getReader();
  const cancel = bestEffortCancel(reader, tracker);
  const decoder = new TextDecoder();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    while (true) {
      const part = await raceReaderRead(reader, signal, cancel, tracker);
      if (part.done) break;
      const value = part.value ?? new Uint8Array();
      total += value.byteLength;
      if (total > maxBytes) throw new Error("Provider response exceeded the bounded body limit.");
      chunks.push(value);
    }
    return chunks.map((chunk) => decoder.decode(chunk, { stream: true })).join("") + decoder.decode();
  } finally {
    cancel();
  }
}

export async function readOpenAIStream(body: ReadableStream<Uint8Array>, callbacks: Pick<AgentCallbacks, "onAssistantDelta">, signal: AbortSignal, trackOperation?: ExternalOperationTracker): Promise<StreamResult> {
  const reader = body.getReader();
  const cancel = bestEffortCancel(reader, trackOperation);
  const decoder = new TextDecoder();
  const calls = new Map<number, ToolCall>();
  let buffer = "";
  let content = "";
  let emitted = false;
  const protocolError = () => Object.assign(new Error("The model stream contained malformed SSE data."), { category: "protocol" });
  const consumeEvent = (event: string): void => {
    const data: string[] = [];
    for (const rawLine of event.split(/\r?\n/)) {
      if (!rawLine || rawLine.startsWith(":")) continue;
      const colon = rawLine.indexOf(":");
      const field = colon < 0 ? rawLine : rawLine.slice(0, colon);
      let value = colon < 0 ? "" : rawLine.slice(colon + 1);
      if (value.startsWith(" ")) value = value.slice(1);
      if (field === "data") data.push(value);
    }
    if (!data.length) return;
    const payload = data.join("\n");
    if (payload.trim() === "[DONE]") return;
    let parsed: { choices?: Array<{ delta?: StreamDelta }> };
    try {
      parsed = JSON.parse(payload) as { choices?: Array<{ delta?: StreamDelta }> };
    } catch {
      throw protocolError();
    }
    const delta = parsed.choices?.[0]?.delta;
    if (!delta) return;
    if (delta.content) {
      content += delta.content;
      emitted = true;
      callbacks.onAssistantDelta(delta.content);
    }
    for (const partial of delta.tool_calls ?? []) {
      emitted = true;
      const index = partial.index ?? 0;
      const current = calls.get(index) ?? { id: partial.id || `call_${index}`, type: "function", function: { name: "", arguments: "" } };
      if (partial.id) current.id = partial.id;
      if (partial.function?.name) current.function.name += partial.function.name;
      if (partial.function?.arguments) current.function.arguments += partial.function.arguments;
      calls.set(index, current);
    }
  };
  const consumeCompleteEvents = (): void => {
    while (true) {
      const boundary = /\r?\n\r?\n/.exec(buffer);
      if (!boundary || boundary.index === undefined) return;
      const event = buffer.slice(0, boundary.index);
      buffer = buffer.slice(boundary.index + boundary[0].length);
      consumeEvent(event);
    }
  };
  try {
    while (true) {
      const part = await raceReaderRead(reader, signal, cancel, trackOperation);
      if (part.done) break;
      buffer += decoder.decode(part.value, { stream: true });
      consumeCompleteEvents();
    }
    buffer += decoder.decode();
    consumeCompleteEvents();
    if (buffer.trim()) throw protocolError();
    return { content, toolCalls: [...calls.values()].filter((call) => call.function.name), emitted };
  } catch (error) {
    if (emitted && !(error instanceof DOMException && (error.name === "AbortError" || error.name === "TimeoutError"))) {
      throw Object.assign(error instanceof Error ? error : protocolError(), { partial: true });
    }
    throw error;
  } finally {
    cancel();
  }
}
interface StreamDelta { content?: string; tool_calls?: Array<{ index?: number; id?: string; function?: { name?: string; arguments?: string } }> }
function parseXmlToolCalls(content: string, runId: string): ToolCall[] { const calls: ToolCall[] = []; const regex = /<tool_call>\s*([\s\S]*?)\s*<\/tool_call>/g; let match: RegExpExecArray | null; let index = 0; while ((match = regex.exec(content))) { const parsed = safeParse(match[1]) as { name?: string; arguments?: unknown } | null; if (parsed?.name) calls.push({ id: generateXmlCallId(runId, index++), type: "function", function: { name: parsed.name, arguments: JSON.stringify(parsed.arguments ?? {}) } }); } return calls; }
function removeXmlToolCalls(content: string): string { return content.replace(/<tool_call>[\s\S]*?<\/tool_call>/g, ""); }
