import { executeExcelTool, executeExcelToolParsed, excelTools, getWorkbookContext, toolSchemas } from "./excelTools";
import { HermesSettings } from "./settings";
import {
  buildServerOwnedRequest,
  isHermesAgentMode,
  isValidBridgeSessionId,
} from "./serverOwnedMode";
import { ActivityItem, ChatMessage, ToolCall, type OpenAITool } from "./types";
import {
  ToolErrorCategory,
  ToolError,
  ToolResult,
  categorizeError,
} from "./toolErrors";
import { validateToolCall } from "./toolValidation";
import {
  isWorkbookPlanProtocolCandidate,
  normalizeWorkbookPlanResponse,
} from "./workbookPlanProtocol";
import type { WorkbookPlanModelResponse } from "./workbookPlanOrchestrator";
import { workbookPlanSchema } from "./generated/workbookPlanSchema";
import type { PrimitiveMutationAuthority } from "./primitiveMutationAuthority";

export type WorkbookPlanModelContinuation = (
  continuation?: string,
  signal?: AbortSignal,
) => Promise<WorkbookPlanModelResponse>;

export interface AgentCallbacks {
  /** Metadata is observed before either successful or failed chat bodies. */
  onResponseHeaders?: (headers: Headers) => void;
  onAssistantStart: () => void;
  onAssistantDelta: (delta: string) => void;
  onAssistantFinal: (content: string) => void;
  onActivity: (item: ActivityItem) => void;
  onStatus: (status: string) => void;
  /**
   * Nonblocking checkpoint. Called at 30s. The agent awaits this Promise.
   * Resolve true = Continue, false = Stop. The turn deadline timer runs
   * concurrently and will reject the checkpoint Promise on timeout.
   */
  onCheckpoint?: (deadlineSignal: AbortSignal) => Promise<boolean>;
  /** Optional route-neutral Task 11 seam. Missing seam always fails closed. */
  onWorkbookPlanTurn?: (
    response: WorkbookPlanModelResponse,
    signal: AbortSignal,
    turnStart: number,
    continueModel: WorkbookPlanModelContinuation,
  ) => Promise<string>;
  /**
   * R4c: the pane turn's primitive mutation authority. Every direct-route
   * primitive executes through it; without one, the read-only sink is used.
   */
  primitiveAuthority?: PrimitiveMutationAuthority;
}

export type StopReason = "checkpoint" | "deadline" | "user" | "timeout";

export interface AgentRunResult {
  messages: ChatMessage[];
  finalContent: string;
  /** Why the turn stopped (undefined = completed normally). */
  stopReason?: StopReason;
}

interface StreamResult {
  content: string;
  toolCalls: ToolCall[];
}

/** Maximum model-call time in ms (20s). */
export const MODEL_CALL_TIMEOUT_MS = 20_000;
/** Maximum total turn time in ms (90s). */
export const TURN_TIMEOUT_MS = 90_000;
/** Checkpoint fires after this many ms of the turn (30s). */
const CHECKPOINT_MS = 30_000;
/** Maximum retry attempts per model call. */
const MAX_RETRIES = 2;
/** Base retry delay in ms. */
const RETRY_BASE_DELAY_MS = 1000;
/** Maximum retry delay cap. */
const RETRY_MAX_DELAY_MS = 8000;
/** Absolute iteration cap regardless of settings. */
const ABSOLUTE_MAX_ITERATIONS = 20;

/** Serialize only the direct agent's parsed/constructed ToolCall records. */
function serializeDirectToolCalls(toolCalls: readonly ToolCall[]): string | undefined {
  if (!Array.isArray(toolCalls)) return undefined;
  try {
    const serialized = JSON.stringify(toolCalls);
    return typeof serialized === "string" ? serialized : undefined;
  } catch {
    return undefined;
  }
}

// Allow tests to inject a sleep implementation (repair 7, model-call containment)
export const sleepHooks = {
  injectedSleep: null as ((ms: number, signal: AbortSignal) => Promise<void>) | null,
};

export function resetInjectedSleep(): void {
  sleepHooks.injectedSleep = null;
}

const TITLE_SPAN_RECIPE = `
Title-span helper (CenterAcrossSelection):
For a title spanning columns, first read the full intended single-row span.
Proceed only when every non-anchor cell is blank and the span is not merged.
Write the title only to the leftmost anchor, then format the full span with
horizontalAlignment CenterAcrossSelection, wrapText false, and columnAutofit
false. No merge capability exists; refuse merge requests and explain.
`;

const SYSTEM_PROMPT = `You are Hermes, an agentic AI assistant running inside the user's active Microsoft Excel workbook.

You can inspect and edit the workbook through tools. Work step by step. Read before writing unless the user explicitly asks you to create new content from scratch. Prefer small, verifiable actions over large risky edits. If a tool returns an error, use that result to correct your next action. When you finish, summarize what you changed and mention any assumptions.

Important tool guidance:
- Always use exact sheet names from workbook context.
- For write_range, the address dimensions must exactly match the 2D values array.
- Strings beginning with "=" are formulas.
- Use format_range and insert_table after writing structured data when useful.
- Use run_calculation after adding or changing formulas.
- Use CenterAcrossSelection (not merge) to span a title across columns.${TITLE_SPAN_RECIPE}

If the endpoint does not support OpenAI tool calls, emit one or more Hermes native calls as XML:
<tool_call>{"name":"tool_name","arguments":{"arg":"value"}}</tool_call>
Only emit XML tool calls when you need a tool.`;

const PLAN_MODE_GUIDANCE = `

Plan mode (enabled only when the WorkbookPlan rollout setting is explicitly true):
- Use exactly one static excel_execute_plan function call for a bounded WorkbookPlan.
- Return one inspect plan for required read-only context, then one execute, ask, or refuse outcome.
- Do not mix excel_execute_plan with primitive calls or prose in the same response.
- The deterministic Excel kernel owns policy, resource limits, ordering, mutation, and verification; do not claim success yourself.`;

export const DIRECT_WORKBOOK_PLAN_TOOL: OpenAITool = Object.freeze({
  type: "function",
  function: {
    name: "excel_execute_plan",
    description: "Submit one bounded WorkbookPlan for deterministic inspect-or-execute handling. Use plan mode guidance and never mix this function with primitive Excel calls.",
    parameters: workbookPlanSchema as unknown as Record<string, unknown>,
  },
});

export function getDirectProviderTools(
  settings: Pick<HermesSettings, "workbookPlanRolloutEnabled">,
): readonly OpenAITool[] {
  return settings.workbookPlanRolloutEnabled === true
    ? [...toolSchemas, DIRECT_WORKBOOK_PLAN_TOOL]
    : [...toolSchemas];
}

/** Exposes the exact direct-mode production prompt for contract tests. */
export function getDirectSystemPrompt(planModeEnabled = false): string {
  return planModeEnabled ? `${SYSTEM_PROMPT}${PLAN_MODE_GUIDANCE}` : SYSTEM_PROMPT;
}

const SERVER_OWNED_SYSTEM_PROMPT = `You are Hermes, operating inside the user's active Microsoft Excel workbook.

The Hermes server owns this entire agent loop. Use the native sidecar_excel_* tools available to you. Those tools are automatically bound to the active Excel workbook session. Execute workbook actions through those native tools and use their results before answering.

Do not emit tool calls as text, JSON, or XML. Do not ask the task pane to execute a tool. When the requested workbook action has completed, return only a concise final summary.
${TITLE_SPAN_RECIPE}`;

/** Exposed for testing: the SERVER_OWNED_SYSTEM_PROMPT. */
export function getServerOwnedSystemPrompt(): string {
  return SERVER_OWNED_SYSTEM_PROMPT;
}

export class HermesAgent {
  private messages: ChatMessage[] = [];

  get history(): ChatMessage[] {
    return [...this.messages];
  }

  reset(): void {
    this.messages = [];
  }

  /**
   * Server-owned mode: sends workbook context to the Hermes server, which runs
   * its own native plugin loop. No caller tools are sent,
   * and no local tool execution happens — we just return the server's answer.
   */
  private async runServerOwned(
    userRequest: string,
    settings: HermesSettings,
    callbacks: AgentCallbacks,
    signal: AbortSignal,
    bridgeSessionId: string,
  ): Promise<AgentRunResult> {
    throwIfAborted(signal);
    const workbookContext = await getWorkbookContext();

    const userMessageContent = `Fresh workbook context:\n${workbookContext}\n\nUser request:\n${userRequest}`;
    const trimmedHistory = this.messages.slice(-12);
    const messages: ChatMessage[] = [
      { role: "system", content: SERVER_OWNED_SYSTEM_PROMPT },
      ...trimmedHistory,
      { role: "user", content: userMessageContent },
    ];

    callbacks.onAssistantStart();
    callbacks.onStatus("Thinking (server-owned)");

    // Make the request WITHOUT tools/tool_choice — server runs its own loop
    const result = await this.chatCompletionServerOwned(
      settings,
      messages,
      signal,
      callbacks,
      bridgeSessionId,
    );

    const finalContent = result.content.trim();
    const assistantMessage: ChatMessage = {
      role: "assistant",
      content: finalContent,
    };
    this.messages = [
      ...this.messages,
      { role: "user", content: userMessageContent },
      assistantMessage,
    ];

    callbacks.onAssistantFinal(finalContent);
    callbacks.onStatus("Ready");
    return { messages: this.history, finalContent };
  }

  async run(
    userRequest: string,
    settings: HermesSettings,
    callbacks: AgentCallbacks,
    signal: AbortSignal,
    bridgeSessionId?: string | null,
  ): Promise<AgentRunResult> {
    // Server-owned mode: hermes-agent model with an active bridge session
    if (isHermesAgentMode(settings.model)) {
      if (!isValidBridgeSessionId(bridgeSessionId)) {
        throw new Error(
          "hermes-agent mode requires an active Excel companion bridge session with a valid UUID. " +
            "Open the Sidecar for Excel task pane and try again.",
        );
      }
      return this.runServerOwned(userRequest, settings, callbacks, signal, bridgeSessionId);
    }

    // Local tool-execution mode: non-hermes-agent models use the existing loop
    const turnStart = Date.now();
    const workbookContext = await getWorkbookContext();
    const turnMessages = this.buildTurnMessages(
      userRequest,
      workbookContext,
      settings.workbookPlanRolloutEnabled === true,
    );
    let workingMessages = [...turnMessages];
    let finalContent = "";
    const toolCallCounts = new Map<string, number>();
    let checkpointFired = false;

    callbacks.onAssistantStart();

    const effectiveMaxIterations = Math.min(
      settings.maxIterations,
      ABSOLUTE_MAX_ITERATIONS,
    );

    for (let iteration = 1; iteration <= effectiveMaxIterations; iteration += 1) {
      throwIfAborted(signal);
      this.checkTurnTimeout(turnStart);
      callbacks.onStatus(`Thinking ${iteration}/${effectiveMaxIterations}`);

      // --- Checkpoint: after CHECKPOINT_MS, ask user non-blockingly ---
      if (!checkpointFired && Date.now() - turnStart >= CHECKPOINT_MS) {
        checkpointFired = true;
        if (callbacks.onCheckpoint) {
          const remainingMs = TURN_TIMEOUT_MS - (Date.now() - turnStart);
          const deadline = deadlineAbortSignal(
            Math.max(remainingMs, 0),
            "Turn exceeded the 90-second absolute time limit.",
          );
          try {
            const shouldContinue = await raceAgainstDeadline(
              callbacks.onCheckpoint(deadline.signal),
              deadline.signal,
            );
            deadline.cleanup();
            if (!shouldContinue) {
              finalContent = "Turn stopped at the 30-second checkpoint.";
              callbacks.onAssistantFinal(finalContent);
              callbacks.onStatus("Stopped at checkpoint");
              return {
                messages: this.history,
                finalContent,
                stopReason: "checkpoint",
              };
            }
          } catch (err) {
            deadline.cleanup();
            // Only classify as deadline if the error is actually a TimeoutError
            // from the deadline signal firing. Other callback errors propagate
            // as internal errors.
            if (
              err instanceof DOMException &&
              err.name === "TimeoutError"
            ) {
              finalContent =
                "Turn exceeded the 90-second absolute time limit while waiting for checkpoint input.";
              callbacks.onAssistantFinal(finalContent);
              callbacks.onStatus("Timeout");
              return {
                messages: this.history,
                finalContent,
                stopReason: "deadline",
              };
            }
            throw err;
          }
        }
      }

      const response = await this.chatCompletionWithRetry(
        settings,
        workingMessages,
        signal,
        callbacks,
        turnStart,
      );

      const serializedToolCalls = serializeDirectToolCalls(response.toolCalls);
      if (serializedToolCalls === undefined) {
        // A direct transport serialization failure fails closed before XML or
        // primitive dispatch can inspect or execute any response data.
        finalContent = "Invalid workbook plan protocol envelope";
        const assistantMessage: ChatMessage = {
          role: "assistant",
          content: finalContent,
        };
        this.messages = [
          ...this.messages,
          this.userHistoryMessage(userRequest, workbookContext),
          assistantMessage,
        ];
        callbacks.onAssistantFinal(finalContent);
        callbacks.onStatus("Ready");
        return { messages: this.history, finalContent };
      }

      if (isWorkbookPlanProtocolCandidate(response.content, serializedToolCalls)) {
        const normalized = await normalizeWorkbookPlanResponse(response.content, serializedToolCalls);
        if (normalized.ok && normalized.kind === "none") {
          // `none` is authoritative primitive/final-content fallthrough.
          // Continue below so benign prose is never converted into an error.
        } else {
          if (normalized.ok === false) {
            finalContent = normalized.error.message;
          } else if (settings.workbookPlanRolloutEnabled !== true) {
            finalContent = "Workbook plan normalized; execution is disabled by the rollout setting.";
          } else if (callbacks.onWorkbookPlanTurn === undefined) {
            finalContent = "Workbook plan execution is unavailable for this trusted turn.";
          } else {
            let initialResponseRecorded = false;
            const continueModel: WorkbookPlanModelContinuation = async (
              continuation,
              continuationSignal = signal,
            ) => {
              if (!initialResponseRecorded) {
                const initialToolCalls = response.toolCalls.length > 0
                  ? response.toolCalls
                  : parseXmlToolCalls(response.content);
                workingMessages = [
                  ...workingMessages,
                  initialToolCalls.length > 0
                    ? {
                        role: "assistant",
                        content: removeXmlToolCalls(response.content).trim() || null,
                        tool_calls: initialToolCalls,
                      }
                    : { role: "assistant", content: response.content },
                ];
                initialResponseRecorded = true;
              }
              if (continuation !== undefined) {
                workingMessages = [
                  ...workingMessages,
                  {
                    role: "user",
                    content: `Workbook plan tool result:\n${continuation}`,
                  },
                ];
              }
              const next = await this.chatCompletionWithRetry(
                settings,
                workingMessages,
                continuationSignal,
                callbacks,
                turnStart,
              );
              return { content: next.content, toolCalls: next.toolCalls };
            };
            finalContent = await callbacks.onWorkbookPlanTurn(
              { content: response.content, toolCalls: response.toolCalls },
              signal,
              turnStart,
              continueModel,
            );
          }
          const assistantMessage: ChatMessage = {
            role: "assistant",
            content: finalContent,
          };
          this.messages = [
            ...this.messages,
            this.userHistoryMessage(userRequest, workbookContext),
            assistantMessage,
          ];
          callbacks.onAssistantFinal(finalContent);
          callbacks.onStatus("Ready");
          return { messages: this.history, finalContent };
        }
      }

      const xmlToolCalls = parseXmlToolCalls(response.content);
      const toolCalls =
        response.toolCalls.length > 0 ? response.toolCalls : xmlToolCalls;

      if (toolCalls.length === 0) {
        finalContent = removeXmlToolCalls(response.content).trim();
        const assistantMessage: ChatMessage = {
          role: "assistant",
          content: finalContent,
        };
        this.messages = [
          ...this.messages,
          this.userHistoryMessage(userRequest, workbookContext),
          assistantMessage,
        ];
        callbacks.onAssistantFinal(finalContent);
        callbacks.onStatus("Ready");
        return { messages: this.history, finalContent };
      }

      const assistantToolMessage: ChatMessage = {
        role: "assistant",
        content: removeXmlToolCalls(response.content).trim() || null,
        tool_calls: toolCalls,
      };
      workingMessages = [...workingMessages, assistantToolMessage];

      for (const toolCall of toolCalls) {
        throwIfAborted(signal);
        this.checkTurnTimeout(turnStart);

        const signature = makeToolCallSignature(toolCall);
        const nextCount = (toolCallCounts.get(signature) ?? 0) + 1;
        toolCallCounts.set(signature, nextCount);
        const blocked = isBlockedRepeatedToolCall(nextCount);
        const label = `${toolCall.function.name} ${summarizeArguments(toolCall.function.arguments)}`;
        callbacks.onActivity({ id: toolCall.id, status: "running", label });

        if (blocked) {
          const blockedResult: ToolError = {
            ok: false,
            error:
              "Repeated identical tool call blocked; use prior result or ask for clarification.",
            category: ToolErrorCategory.Repetition,
          };
          callbacks.onActivity({
            id: toolCall.id,
            status: "error",
            label,
            detail: JSON.stringify(blockedResult),
          });
          workingMessages = [
            ...workingMessages,
            {
              role: "tool",
              tool_call_id: toolCall.id,
              name: toolCall.function.name,
              content: JSON.stringify(blockedResult),
            },
          ];
          continue;
        }

        // --- Pre-execution validation ---
        const validation = validateToolCall(
          toolCall.function.name,
          toolCall.function.arguments,
        );

        if ("category" in validation) {
          callbacks.onActivity({
            id: toolCall.id,
            status: "error",
            label,
            detail: JSON.stringify(validation),
          });
          workingMessages = [
            ...workingMessages,
            {
              role: "tool",
              tool_call_id: toolCall.id,
              name: toolCall.function.name,
              content: JSON.stringify(validation),
            },
          ];
          continue;
        }

        // Execute with validated parsed args — never re-parse (repair 1)
        const remainingToolMs = TURN_TIMEOUT_MS - (Date.now() - turnStart);
        let result: string;
        try {
          if (remainingToolMs <= 0) {
            this.checkTurnTimeout(turnStart);
          }
          result = await raceWithTimeout(
            callbacks.primitiveAuthority !== undefined
              ? callbacks.primitiveAuthority.executePrimitive(toolCall.function.name, validation.parsedArgs)
              : executeExcelToolParsed(toolCall.function.name, validation.parsedArgs),
            Math.max(remainingToolMs, 100),
          );
        } catch (toolErr) {
          if (
            toolErr instanceof DOMException &&
            toolErr.name === "TimeoutError"
          ) {
            throw toolErr;
          }
          result = JSON.stringify({
            ok: false,
            error:
              toolErr instanceof Error ? toolErr.message : String(toolErr),
            category: categorizeError(toolErr),
            name: toolErr instanceof DOMException ? toolErr.name : undefined,
          });
        }

        const parsed = safeParseJson(result) as ToolResult | null;
        const isError =
          parsed !== null &&
          typeof parsed === "object" &&
          "ok" in parsed &&
          parsed.ok === false;

        callbacks.onActivity({
          id: toolCall.id,
          status: isError ? "error" : "done",
          label,
          detail: result,
        });
        workingMessages = [
          ...workingMessages,
          {
            role: "tool",
            tool_call_id: toolCall.id,
            name: toolCall.function.name,
            content: result,
          },
        ];
      }
    }

    finalContent = `I reached the max of ${effectiveMaxIterations} agent iterations before finishing. I stopped to avoid making unbounded workbook changes.`;
    callbacks.onAssistantFinal(finalContent);
    callbacks.onStatus("Stopped at iteration limit");
    this.messages = [
      ...this.messages,
      this.userHistoryMessage(userRequest, workbookContext),
      { role: "assistant", content: finalContent },
    ];
    return { messages: this.history, finalContent };
  }

  private buildTurnMessages(
    userRequest: string,
    workbookContext: string,
    planModeEnabled: boolean,
  ): ChatMessage[] {
    const trimmedHistory = this.messages.slice(-12);
    return [
      { role: "system", content: getDirectSystemPrompt(planModeEnabled) },
      ...trimmedHistory,
      this.userHistoryMessage(userRequest, workbookContext),
    ];
  }

  private userHistoryMessage(
    userRequest: string,
    workbookContext: string,
  ): ChatMessage {
    return {
      role: "user",
      content: `Fresh workbook context:\n${workbookContext}\n\nUser request:\n${userRequest}`,
    };
  }

  /**
   * Server-owned chat completion: sends no tools/tool_choice. The Hermes server
   * runs its own native plugin loop and returns the final answer directly.
   */
  private async chatCompletionServerOwned(
    settings: HermesSettings,
    messages: ChatMessage[],
    signal: AbortSignal,
    callbacks: AgentCallbacks,
    bridgeSessionId: string,
  ): Promise<StreamResult> {
    const url = chatCompletionsUrl(settings.baseUrl);
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      ...(settings.requestHeaders ?? {}),
    };
    if (settings.apiKey.trim()) {
      headers.Authorization = `Bearer ${settings.apiKey.trim()}`;
    }
    headers["X-Hermes-Session-Id"] = bridgeSessionId;

    let response: Response;
    try {
      response = await fetch(url, {
        method: "POST",
        headers,
        signal,
        body: JSON.stringify(
          buildServerOwnedRequest(settings, messages, bridgeSessionId),
        ),
      });
    } catch (error) {
      if (signal.aborted) {
        throw new DOMException(
          "The running Hermes task was stopped.",
          "AbortError",
        );
      }
      throw new Error(
        `Could not reach ${settings.baseUrl}. ${error instanceof Error ? error.message : String(error)}`,
      );
    }

    callbacks.onResponseHeaders?.(response.headers);
    if (!response.ok) {
      const errMsg = `Model endpoint returned ${response.status}.`;
      const err = new Error(errMsg);
      (err as unknown as Record<string, unknown>).statusCode = response.status;
      throw err;
    }

    return readOpenAIStream(response.body!, callbacks, signal);
  }

  private async chatCompletionWithRetry(
    settings: HermesSettings,
    messages: ChatMessage[],
    signal: AbortSignal,
    callbacks: AgentCallbacks,
    turnStart: number,
  ): Promise<StreamResult> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      throwIfAborted(signal);
      this.checkTurnTimeout(turnStart);

      try {
        return await this.chatCompletion(
          settings,
          messages,
          signal,
          callbacks,
          turnStart,
        );
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        const category = categorizeError(error);
        if (
          category !== ToolErrorCategory.Network &&
          category !== ToolErrorCategory.Provider
        ) {
          throw lastError;
        }

        if (category === ToolErrorCategory.Provider) {
          const statusCode = extractStatusCode(error);
          if (
            statusCode !== undefined &&
            statusCode >= 400 &&
            statusCode < 500 &&
            statusCode !== 429
          ) {
            throw lastError;
          }
        }

        if (attempt === MAX_RETRIES) {
          throw lastError;
        }

        const remainingTurn = TURN_TIMEOUT_MS - (Date.now() - turnStart);
        const maxDelay = Math.max(0, remainingTurn - 500);
        const rawDelay = Math.min(
          RETRY_BASE_DELAY_MS * Math.pow(2, attempt),
          RETRY_MAX_DELAY_MS,
        );
        const delay = Math.min(rawDelay, maxDelay);

        if (delay <= 0) {
          throw lastError;
        }

        callbacks.onStatus(
          `Retrying in ${(delay / 1000).toFixed(1)}s (attempt ${attempt + 1}/${MAX_RETRIES})`,
        );
        // Use injected sleep if available, otherwise real sleep
        const sleepFn = sleepHooks.injectedSleep ?? interruptibleSleep;
        await sleepFn(delay, signal);
      }
    }

    throw lastError ?? new Error("chatCompletionWithRetry: unreachable");
  }

  private async chatCompletion(
    settings: HermesSettings,
    messages: ChatMessage[],
    signal: AbortSignal,
    _callbacks: AgentCallbacks,
    turnStart: number,
  ): Promise<StreamResult> {
    const url = chatCompletionsUrl(settings.baseUrl);
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      ...(settings.requestHeaders ?? {}),
    };
    if (settings.apiKey.trim()) {
      headers.Authorization = `Bearer ${settings.apiKey.trim()}`;
    }

    // Per-call budget: 20s max, capped by remaining 90s turn budget
    const remainingTurn = TURN_TIMEOUT_MS - (Date.now() - turnStart);
    const perCallTimeout = Math.min(
      MODEL_CALL_TIMEOUT_MS,
      Math.max(remainingTurn, 1000),
    );

    const abortController = new AbortController();
    const linked = linkSignalsCleanup(signal, abortController.signal);

    // Single timeoutId lives for the ENTIRE call: fetch + body + parse.
    // It is ONLY cleared in the finally block.
    const timeoutId = setTimeout(() => {
      abortController.abort(
        new DOMException("Per-call timeout", "TimeoutError"),
      );
    }, perCallTimeout);

    try {
      let response: Response;
      try {
        response = await fetch(url, {
          method: "POST",
          headers,
          signal: linked.signal,
          body: JSON.stringify({
            model: settings.model,
            messages,
            tools: getDirectProviderTools(settings),
            tool_choice: "auto",
            stream: true,
            temperature: 0.2,
          }),
        });
      } catch (error) {
        // User cancellation first — clean AbortError, not network error
        if (signal.aborted) {
          throw new DOMException(
            "The running Hermes task was stopped.",
            "AbortError",
          );
        }
        if (linked.signal.aborted) {
          // Per-call timeout (only remaining possibility since user signal is not aborted)
          throw new DOMException(
            `Model call timed out after ${perCallTimeout / 1000}s.`,
            "TimeoutError",
          );
        }
        // Genuine network/fetch failure
        throw new Error(
          `Could not reach ${settings.baseUrl}. ${error instanceof Error ? error.message : String(error)}`,
        );
      }

      _callbacks.onResponseHeaders?.(response.headers);
      if (!response.ok) {
        // Sanitized: only log status, never append raw provider body
        const errMsg = `Model endpoint returned ${response.status}.`;
        const err = new Error(errMsg);
        (err as unknown as Record<string, unknown>).statusCode = response.status;
        throw err;
      }

      const contentType = response.headers.get("content-type") || "";

      // --- Non-stream JSON fallback on 2xx ---
      if (!response.body || contentType.includes("application/json")) {
        if (!response.body) {
          return { content: "", toolCalls: [] };
        }
        const jsonText = await readResponseBody(
          response.body,
          linked.signal,
          perCallTimeout,
        );
        const parsed = safeParseJson(jsonText) as {
          choices?: Array<{
            message?: { content?: string; tool_calls?: ToolCall[] };
          }>;
        } | null;
        if (parsed?.choices?.[0]?.message) {
          const msg = parsed.choices[0].message;
          return {
            content: msg.content ?? "",
            toolCalls: msg.tool_calls ?? [],
          };
        }
        return { content: "", toolCalls: [] };
      }

      // --- Streaming path ---
      const result = await readOpenAIStream(
        response.body,
        _callbacks,
        linked.signal,
      );
      return result;
    } finally {
      // Single finalization: clear the per-call timer and signal listeners
      clearTimeout(timeoutId);
      linked.cleanup();
    }
  }

  private checkTurnTimeout(turnStart: number): void {
    if (Date.now() - turnStart >= TURN_TIMEOUT_MS) {
      throw new DOMException(
        "Turn exceeded the 90-second absolute time limit.",
        "TimeoutError",
      );
    }
  }
}

// ---------------------------------------------------------------------------
// Streaming parser (exported for tests)
// ---------------------------------------------------------------------------

export async function readOpenAIStream(
  body: ReadableStream<Uint8Array>,
  callbacks: AgentCallbacks,
  signal: AbortSignal,
): Promise<StreamResult> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  const toolCallsByIndex = new Map<number, ToolCall>();
  let buffer = "";
  let content = "";

  // Cancel the reader when signal fires, so pending reader.read() resolves
  let aborted = false;
  const onAbort = () => {
    aborted = true;
    reader.cancel().catch(() => {});
  };
  signal.addEventListener("abort", onAbort, { once: true });

  const consumeEvent = (event: string) => {
    const lines = event
      .split("\n")
      .map((line) => line.trim())
      .filter((line) => line.startsWith("data:"));
    for (const line of lines) {
      const payload = line.replace(/^data:\s*/, "");
      if (payload === "[DONE]") {
        continue;
      }

      const parsed = safeParseJson(payload) as {
        choices?: Array<{
          delta?: StreamDelta;
          message?: { content?: string | null; tool_calls?: ToolCall[] };
          finish_reason?: string;
        }>;
      } | null;
      const choice = parsed?.choices?.[0];
      if (!choice) {
        continue;
      }

      // --- SSE-inline one-shot fallback (repair 4) ---
      // Some servers emit a full message inside an SSE data event instead of
      // incremental deltas. Handle this without duplicating normal delta output.
      if (choice.message) {
        if (typeof choice.message.content === "string") {
          content = mergeStreamingChunk(content, choice.message.content);
          callbacks.onAssistantDelta(choice.message.content);
        }
        if (choice.message.tool_calls) {
          for (const tc of choice.message.tool_calls) {
            const key = toolCallsByIndex.size;
            const existing: ToolCall = {
              id: tc.id || `call_${Date.now()}_${key}`,
              type: "function",
              function: { name: "", arguments: "" },
            };
            if (tc.id) existing.id = tc.id;
            if (tc.function?.name) {
              existing.function.name = mergeStreamingChunk(
                existing.function.name,
                tc.function.name,
              );
            }
            if (tc.function?.arguments) {
              existing.function.arguments = mergeStreamingChunk(
                existing.function.arguments,
                tc.function.arguments,
              );
            }
            toolCallsByIndex.set(key, existing);
          }
        }
        continue;
      }

      const delta = choice.delta;
      if (!delta) {
        continue;
      }

      if (delta.content) {
        content = mergeStreamingChunk(content, delta.content);
        callbacks.onAssistantDelta(delta.content);
      }

      if (delta.tool_calls) {
        for (const partial of delta.tool_calls) {
          const index = partial.index ?? 0;
          const existing = toolCallsByIndex.get(index) || {
            id: partial.id || `call_${Date.now()}_${index}`,
            type: "function" as const,
            function: { name: "", arguments: "" },
          };
          if (partial.id) {
            existing.id = partial.id;
          }
          if (partial.function?.name) {
            existing.function.name = mergeStreamingChunk(
              existing.function.name,
              partial.function.name,
            );
          }
          if (partial.function?.arguments) {
            existing.function.arguments = mergeStreamingChunk(
              existing.function.arguments,
              partial.function.arguments,
            );
          }
          toolCallsByIndex.set(index, existing);
        }
      }
    }
  };

  while (true) {
    throwIfAborted(signal);
    const { value, done } = await reader.read();
    const decoded = value ? decoder.decode(value, { stream: true }) : "";
    buffer += decoded;

    const events = buffer.split("\n\n");
    buffer = events.pop() || "";

    for (const event of events) {
      consumeEvent(event);
    }

    if (done) {
      if (buffer.trim()) {
        consumeEvent(buffer);
        buffer = "";
      }
      break;
    }
  }

  // Clean up the reader-cancel listener on normal completion
  signal.removeEventListener("abort", onAbort);

  // If the signal aborted during streaming, re-throw rather than
  // returning empty content from a cancelled reader
  if (aborted) {
    const reason = signal.reason;
    if (reason instanceof DOMException) {
      throw reason;
    }
    throw new DOMException("Model call aborted.", "AbortError");
  }

  return {
    content,
    toolCalls: [...toolCallsByIndex.values()].filter(
      (call) => call.function.name,
    ),
  };
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

function mergeStreamingChunk(base: string, incoming: string): string {
  return `${base}${incoming}`;
}

interface StreamDelta {
  content?: string;
  tool_calls?: Array<{
    index?: number;
    id?: string;
    type?: "function";
    function?: {
      name?: string;
      arguments?: string;
    };
  }>;
}

export function chatCompletionsUrl(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/, "");
  if (!trimmed) {
    throw new Error("Hermes is not connected. Click Connect to Hermes first.");
  }
  if (trimmed.endsWith("/chat/completions")) return trimmed;
  return trimmed.endsWith("/v1") ? `${trimmed}/chat/completions` : `${trimmed}/v1/chat/completions`;
}

function parseXmlToolCalls(content: string): ToolCall[] {
  const calls: ToolCall[] = [];
  const regex = /<tool_call>\s*([\s\S]*?)\s*<\/tool_call>/g;
  let match: RegExpExecArray | null;
  let index = 0;

  while ((match = regex.exec(content)) !== null) {
    const parsed = safeParseJson(match[1]) as {
      name?: string;
      arguments?: unknown;
    } | null;
    if (!parsed?.name) {
      continue;
    }
    calls.push({
      id: `xml_call_${Date.now()}_${index}`,
      type: "function",
      function: {
        name: parsed.name,
        arguments: JSON.stringify(parsed.arguments || {}),
      },
    });
    index += 1;
  }

  return calls;
}

function removeXmlToolCalls(content: string): string {
  return content.replace(/<tool_call>[\s\S]*?<\/tool_call>/g, "");
}

function summarizeArguments(rawArgs: string): string {
  const parsed = safeParseJson(rawArgs);
  if (!parsed || typeof parsed !== "object") {
    return "";
  }
  const args = parsed as Record<string, unknown>;
  const sheet =
    typeof args.sheet === "string"
      ? args.sheet
      : typeof args.name === "string"
        ? args.name
        : "";
  const address =
    typeof args.address === "string"
      ? args.address
      : typeof args.dataRange === "string"
        ? args.dataRange
        : "";
  return [sheet, address].filter(Boolean).join("!");
}

function safeParseJson(value: string): unknown | null {
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

export function makeToolCallSignature(toolCall: ToolCall): string {
  return `${toolCall.function.name}:${canonicalizeToolArguments(toolCall.function.arguments)}`;
}

function canonicalizeToolArguments(rawArguments: string): string {
  const parsed = safeParseJson(rawArguments);
  if (parsed === null) {
    return rawArguments;
  }
  return JSON.stringify(canonicalizeJsonValue(parsed));
}

function canonicalizeJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => canonicalizeJsonValue(item));
  }

  if (value && typeof value === "object") {
    const entries = Object.entries(value)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([key, nested]) => [key, canonicalizeJsonValue(nested)] as const);

    return Object.fromEntries(entries);
  }

  return value;
}

export function isBlockedRepeatedToolCall(count: number): boolean {
  return count > 2;
}

export function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) {
    throw new DOMException("The running Hermes task was stopped.", "AbortError");
  }
}

// ---------------------------------------------------------------------------
// Signal linking with cleanup
// ---------------------------------------------------------------------------

interface LinkedSignalResult {
  signal: AbortSignal;
  cleanup: () => void;
}

function linkSignalsCleanup(...signals: AbortSignal[]): LinkedSignalResult {
  if (typeof AbortSignal.any === "function") {
    const signal = AbortSignal.any(signals);
    return { signal, cleanup: () => {} };
  }

  const controller = new AbortController();
  const listeners: Array<{ signal: AbortSignal; handler: () => void }> = [];

  const onAbort = () => {
    controller.abort();
  };

  for (const sig of signals) {
    if (sig.aborted) {
      controller.abort();
      break;
    }
    sig.addEventListener("abort", onAbort, { once: true });
    listeners.push({ signal: sig, handler: onAbort });
  }

  return {
    signal: controller.signal,
    cleanup: () => {
      for (const { signal: sig, handler } of listeners) {
        sig.removeEventListener("abort", handler);
      }
      listeners.length = 0;
    },
  };
}

// ---------------------------------------------------------------------------
// Deadline & timeout helpers
// ---------------------------------------------------------------------------

export interface DeadlineSignal {
  signal: AbortSignal;
  /** Clear the underlying timer. Call when the deadline resolves early. */
  cleanup: () => void;
}

export interface TurnCancellation {
  readonly signal: AbortSignal;
  readonly cleanup: () => void;
}

/**
 * Create an AbortSignal that fires after `ms` milliseconds.
 * Returns a cleanup function to clear the timer if the deadline resolves
 * before firing (repair 5).
 */
export function deadlineAbortSignal(
  ms: number,
  reason: string,
): DeadlineSignal {
  const controller = new AbortController();
  let timerId: ReturnType<typeof setTimeout> | null = null;
  if (ms <= 0) {
    controller.abort(new DOMException(reason, "TimeoutError"));
  } else {
    timerId = setTimeout(() => {
      controller.abort(new DOMException(reason, "TimeoutError"));
    }, ms);
  }
  return {
    signal: controller.signal,
    cleanup: () => {
      if (timerId !== null) {
        clearTimeout(timerId);
        timerId = null;
      }
    },
  };
}

export function createTurnCancellation(
  turnStart: number,
  userSignal: AbortSignal,
): TurnCancellation {
  const remainingMs = TURN_TIMEOUT_MS - (Date.now() - turnStart);
  const deadline = deadlineAbortSignal(
    remainingMs,
    "Turn exceeded the 90-second absolute time limit.",
  );
  const controller = new AbortController();
  const abortFromUser = () => {
    if (!controller.signal.aborted) {
      controller.abort(
        userSignal.reason ?? new DOMException("The running Hermes task was stopped.", "AbortError"),
      );
    }
  };
  const abortFromDeadline = () => {
    if (!controller.signal.aborted) {
      controller.abort(
        deadline.signal.reason ?? new DOMException("Turn exceeded the 90-second absolute time limit.", "TimeoutError"),
      );
    }
  };
  if (userSignal.aborted) abortFromUser();
  if (deadline.signal.aborted) abortFromDeadline();
  userSignal.addEventListener("abort", abortFromUser, { once: true });
  deadline.signal.addEventListener("abort", abortFromDeadline, { once: true });
  return {
    signal: controller.signal,
    cleanup: () => {
      userSignal.removeEventListener("abort", abortFromUser);
      deadline.signal.removeEventListener("abort", abortFromDeadline);
      deadline.cleanup();
    },
  };
}

async function raceAgainstDeadline<T>(
  promise: Promise<T>,
  signal: AbortSignal,
): Promise<T> {
  if (signal.aborted) {
    throw signal.reason ?? new DOMException("Deadline reached.", "TimeoutError");
  }

  return new Promise<T>((resolve, reject) => {
    const onAbort = () => {
      reject(signal.reason ?? new DOMException("Deadline reached.", "TimeoutError"));
    };
    signal.addEventListener("abort", onAbort, { once: true });

    promise.then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (err) => {
        signal.removeEventListener("abort", onAbort);
        reject(err);
      },
    );
  });
}

function raceWithTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
): Promise<T> {
  if (timeoutMs <= 0) {
    return Promise.reject(
      new DOMException("Operation timed out.", "TimeoutError"),
    );
  }

  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new DOMException("Operation timed out.", "TimeoutError"));
    }, timeoutMs);

    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      },
    );
  });
}

export function interruptibleSleep(
  ms: number,
  signal: AbortSignal,
): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    if (signal.aborted) {
      reject(new DOMException("Aborted during retry delay.", "AbortError"));
      return;
    }
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    const onAbort = () => {
      clearTimeout(timer);
      reject(new DOMException("Aborted during retry delay.", "AbortError"));
    };
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

// ---------------------------------------------------------------------------
// HTTP helpers
// ---------------------------------------------------------------------------

async function readResponseBody(
  body: ReadableStream<Uint8Array>,
  signal: AbortSignal,
  timeoutMs: number,
): Promise<string> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  let result = "";

  let timerId: ReturnType<typeof setTimeout> | null = null;
  let aborted = false;
  let cancelled = false;

  // Cancel the reader when signal fires or on explicit cleanup
  const cancelReader = () => {
    if (cancelled) return;
    cancelled = true;
    reader.cancel().catch(() => {});
  };

  // Cancel the reader and clear timer when signal fires
  const onAbort = () => {
    aborted = true;
    if (timerId !== null) {
      clearTimeout(timerId);
      timerId = null;
    }
    cancelReader();
  };
  signal.addEventListener("abort", onAbort, { once: true });

  // Shared cleanup: clear timer, remove listener, cancel reader
  const cleanup = () => {
    if (timerId !== null) {
      clearTimeout(timerId);
      timerId = null;
    }
    signal.removeEventListener("abort", onAbort);
    cancelReader();
  };

  const timeoutPromise = new Promise<string>((_, reject) => {
    timerId = setTimeout(() => {
      cleanup();
      reject(new DOMException("Body read timed out.", "TimeoutError"));
    }, timeoutMs);
  });

  const readPromise = (async () => {
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (value) result += decoder.decode(value, { stream: true });
        if (done) {
          result += decoder.decode();
          return result;
        }
      }
    } finally {
      cleanup();
      // If signal aborted, re-throw with correct error type
      if (aborted) {
        const reason = signal.reason;
        if (reason instanceof DOMException) {
          throw reason;
        }
        throw new DOMException("Body read aborted.", "AbortError");
      }
    }
  })();

  return Promise.race([readPromise, timeoutPromise]);
}

function extractStatusCode(error: unknown): number | undefined {
  if (typeof error === "object" && error !== null) {
    const record = error as Record<string, unknown>;
    if (typeof record.statusCode === "number") return record.statusCode;
    if (typeof record.status === "number") return record.status;
  }
  return undefined;
}
