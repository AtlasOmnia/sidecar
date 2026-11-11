import { ApprovalRequiredError, GRANT_HEADER } from "./deskConnect";
import { WORKBOOK_PLAN_FINGERPRINT } from "./generated/workbookPlanSchema";
import { WORKBOOK_PLAN_TOOL_NAME } from "./workbookPlanProtocol";
import { SERVER_INSPECTION_SUPPORTED } from "./taskpane/serverOwnedPlanRoute";

const DEFAULT_BASE_URL = "";
const PROTOCOL_VERSION = "1.0";
const HEARTBEAT_MS = 15_000;
const POLL_MS = 250;
const INTERNAL_PLAN_TOOL = WORKBOOK_PLAN_TOOL_NAME;

/**
 * Fixed bound for the one-shot POST /excel/execute-start acknowledgement.
 * The pane cannot see the remaining claim lease, so this is a constant chosen
 * well below the companion's 20 s CLAIM_LEASE_SECONDS: a late or lost
 * acknowledgement is treated as a rejection (zero mutation), and the
 * companion's own sweep stays the authority for the ticket's final state.
 */
export const EXECUTE_START_ACK_TIMEOUT_MS = 5_000;

const ACK_REJECTED_MESSAGE = "Workbook plan execution start was not acknowledged.";
const INVALID_PLAN_TICKET_MESSAGE = "Workbook plan ticket was invalid and was not executed.";
const PLAN_NOT_EXECUTED_MESSAGE = "Workbook plan was not executed.";
const PLAN_UNCONFIRMED_MESSAGE = "Workbook plan execution ended without a confirmed result.";
/** At most this many result POSTs per ticket (the initial post plus two re-posts), then the entry is dropped. */
const MAX_RESULT_POST_ATTEMPTS = 3;
const OVERSIZE_PRIMITIVE_RESULT_MESSAGE = "Office tool result was too large to submit.";
/** Mirrors scripts/companion/protocol.py MAX_RESULT_BODY: a transport failure on a larger result post is treated as oversize. */
const MAX_RESULT_BODY_BYTES = 10 * 1024 * 1024;
const MAX_RESPONSE_ENVELOPE_BYTES = 64 * 1024;
const MAX_RESPONSE_BYTES = MAX_RESULT_BODY_BYTES + MAX_RESPONSE_ENVELOPE_BYTES;
/** Per bridge session, the last this-many dispatched ticket ids are never handed to executeTool again (FIFO). */
const MAX_DISPATCHED_TICKETS_PER_SESSION = 32;

export interface BridgePublicState {
  connected: boolean;
  sessionId: string | null;
  documentId: string | null;
  lastError?: string;
}

export interface BridgeSession {
  sessionId: string;
  documentId: string;
  protocolVersion: string;
  expiresInSeconds: number;
}

interface PrivateSession extends BridgeSession {
  sessionKey: string;
  workbookPlanFingerprint: string | null;
}

export interface OfficeBridgeDependencies {
  hostPrefix: string;
  documentId: () => Promise<string>;
  paneId: string;
  executeTool: (
    name: string,
    argumentsValue: Record<string, unknown>,
    trusted?: BridgeTrustedPlanEnvelope,
    execution?: BridgePlanExecutionControl,
  ) => Promise<string>;
  fetchImpl?: typeof fetch;
  baseUrl?: string;
  onState?: (state: BridgePublicState) => void;
  grant?: () => string | null;
  onApprovalRequired?: () => void;
  setIntervalImpl?: typeof setInterval;
  clearIntervalImpl?: typeof clearInterval;
}

export interface BridgeTrustedPlanEnvelope {
  turnOperationId: string;
  workbookPlanFingerprint: typeof WORKBOOK_PLAN_FINGERPRINT;
}

/**
 * R4g-2a: the v2 trusted plan envelope the companion delivers on every plan
 * ticket (exactly these four keys). `turnKey` names the Hermes turn shared by
 * its inspect and execute operations and never equals `turnOperationId`.
 */
export interface BridgeTrustedPlanEnvelopeV2 extends BridgeTrustedPlanEnvelope {
  readonly phase: "inspect" | "execute";
  readonly turnKey: string;
}

/**
 * Ticket-bound execution control handed to the pane for trusted plan tickets
 * only. `acknowledgeExecuteStart` is one-shot: it POSTs /excel/execute-start at
 * most once and always settles (fixed timeout, optional turn AbortSignal).
 */
export interface BridgePlanExecutionControl {
  readonly acknowledgeExecuteStart: (signal?: AbortSignal) => Promise<void>;
}

interface CompletedResult {
  readonly result: unknown;
  /** Present for plan tickets only: whether the one-shot acknowledgement resolved. */
  readonly executed?: boolean;
  /** Result POSTs attempted so far for this ticket (mutated in place). */
  attempts?: number;
  /** Set after a 413: later attempts post the fixed content-free body. */
  oversize?: boolean;
  /** The session whose poll produced this entry; never serialized. */
  readonly owner?: PrivateSession;
}

class BridgeHttpError extends Error {
  constructor(public readonly status: number, message: string) {
    super(message);
    this.name = "BridgeHttpError";
  }
}

export class OfficeCompanionBridge {
  private readonly fetchImpl: typeof fetch;
  private readonly baseUrl: string;
  private readonly setIntervalImpl: typeof setInterval;
  private readonly clearIntervalImpl: typeof clearInterval;
  private session: PrivateSession | null = null;
  private connecting: Promise<BridgeSession> | null = null;
  private readonly completedResults = new Map<string, CompletedResult>();
  private readonly dispatchedTickets = new WeakMap<PrivateSession, string[]>();
  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;
  private pollTimer: ReturnType<typeof setInterval> | null = null;
  /** R4j (R4f lane-B B4): the (session, generation) whose timer-started poll is in flight; replaces the bridge-wide flag. */
  private timerPoll: { readonly session: PrivateSession | null; readonly generation: number } | null = null;
  /** R4j: executeTool calls in flight anywhere on this bridge; the timer starts no poll while it is non-zero. */
  private activeToolCalls = 0;
  private running = false;
  private lifecycleGeneration = 0;

  constructor(private readonly deps: OfficeBridgeDependencies) {
    const fetchImpl = deps.fetchImpl ?? globalThis.fetch;
    this.fetchImpl = fetchImpl.bind(globalThis);
    this.baseUrl = (deps.baseUrl ?? DEFAULT_BASE_URL).replace(/\/+$/, "");
    const setIntervalImpl = deps.setIntervalImpl ?? globalThis.setInterval;
    const clearIntervalImpl = deps.clearIntervalImpl ?? globalThis.clearInterval;
    this.setIntervalImpl = setIntervalImpl.bind(globalThis);
    this.clearIntervalImpl = clearIntervalImpl.bind(globalThis);
    if (!/^[a-z][a-z0-9-]{0,31}$/.test(deps.hostPrefix)) {
      throw new Error("Invalid bridge host prefix");
    }
    if (!isUuid(deps.paneId)) {
      throw new Error("paneId must be a UUID");
    }
  }

  get sessionId(): string | null {
    return this.session?.sessionId ?? null;
  }

  get documentId(): string | null {
    return this.session?.documentId ?? null;
  }

  get publicState(): BridgePublicState {
    return {
      connected: this.session !== null,
      sessionId: this.sessionId,
      documentId: this.documentId,
    };
  }

  connect(): Promise<BridgeSession> {
    if (this.session) {
      return Promise.resolve(this.publicSession(this.session));
    }
    if (this.connecting) {
      return this.connecting;
    }
    const connecting = this.connectNewSession();
    this.connecting = connecting;
    connecting.then(
      () => { if (this.connecting === connecting) this.connecting = null; },
      () => { if (this.connecting === connecting) this.connecting = null; },
    );
    return connecting;
  }

  private async connectNewSession(): Promise<BridgeSession> {
    const documentId = await this.deps.documentId();
    if (!documentId || documentId.length > 256) {
      throw new Error("documentId must be a bounded non-empty string");
    }
    const grant = this.deps.grant?.() ?? null;
    const response = await this.fetchImpl(this.url("session"), {
      method: "POST",
      headers: { "Content-Type": "application/json", ...(grant ? { [GRANT_HEADER]: grant } : {}) },
      body: JSON.stringify({
        protocolVersion: PROTOCOL_VERSION,
        documentId,
        paneId: this.deps.paneId,
        workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
      }),
    });
    const value = await readJsonObject(response);
    if (response.status === 403 && value.approval === "required") {
      this.deps.onApprovalRequired?.();
      throw new ApprovalRequiredError();
    }
    if (response.status !== 201) {
      throw new BridgeHttpError(response.status, responseMessage(value, "Session registration failed"));
    }
    const sessionId = value.sessionId;
    const sessionKey = value.sessionKey;
    const protocolVersion = value.protocolVersion;
    const expiresInSeconds = value.expiresInSeconds;
    if (!isUuid(sessionId) || typeof sessionKey !== "string" || sessionKey.length < 32) {
      throw new Error("Companion returned an invalid session identity");
    }
    if (protocolVersion !== PROTOCOL_VERSION || typeof expiresInSeconds !== "number") {
      throw new Error("Companion returned an incompatible protocol response");
    }
    const negotiatedFingerprint = value.workbookPlanFingerprint;
    if (negotiatedFingerprint !== undefined
        && negotiatedFingerprint !== WORKBOOK_PLAN_FINGERPRINT) {
      throw new Error("Companion returned an incompatible WorkbookPlan fingerprint");
    }
    this.session = {
      sessionId,
      sessionKey,
      documentId,
      protocolVersion,
      expiresInSeconds,
      workbookPlanFingerprint: negotiatedFingerprint === WORKBOOK_PLAN_FINGERPRINT
        ? WORKBOOK_PLAN_FINGERPRINT
        : null,
    };
    this.emitState();
    return this.publicSession(this.session);
  }

  async start(): Promise<BridgeSession> {
    const generation = this.lifecycleGeneration;
    const session = await this.connect();
    if (generation !== this.lifecycleGeneration) {
      return session;
    }
    if (this.running) {
      return session;
    }
    this.running = true;
    this.heartbeatTimer = this.setIntervalImpl(() => {
      const session = this.session;
      const generation = this.lifecycleGeneration;
      void this.heartbeatOnce().catch((error) => this.handleBackgroundError(error, session, generation));
    }, HEARTBEAT_MS);
    this.pollTimer = this.setIntervalImpl(() => {
      const session = this.session;
      const generation = this.lifecycleGeneration;
      if (this.activeToolCalls > 0) return;
      const running = this.timerPoll;
      if (running !== null && running.session === session && running.generation === generation) return;
      const token = { session, generation };
      this.timerPoll = token;
      void this.pollOnce()
        .catch((error) => this.handleBackgroundError(error, session, generation))
        .finally(() => { if (this.timerPoll === token) this.timerPoll = null; });
    }, POLL_MS);
    return session;
  }

  async heartbeatOnce(): Promise<void> {
    const session = this.requireSession();
    const response = await this.fetchImpl(this.url("heartbeat"), {
      method: "POST",
      headers: this.sessionHeaders(session, true),
      body: "{}",
    });
    if (response.status !== 200) {
      const value = await readJsonObject(response);
      throw new BridgeHttpError(response.status, responseMessage(value, "Heartbeat failed"));
    }
  }

  async pollOnce(): Promise<boolean> {
    const session = this.requireSession();
    const generation = this.lifecycleGeneration;
    // R4j (R4f lane-B B2): an entry from a replaced session is dropped unposted; it is never re-posted
    // with this session's credentials.
    for (const [cachedTicketId, entry] of this.completedResults) if (entry.owner !== session) this.completedResults.delete(cachedTicketId);
    const cached = this.completedResults.entries().next();
    if (!cached.done) {
      const [ticketId, completed] = cached.value;
      await this.flushCompletedResult(session, ticketId, completed);
      return true;
    }
    const response = await this.fetchImpl(this.url("pending"), {
      method: "GET",
      headers: this.sessionHeaders(session, false),
    });
    if (response.status === 204) {
      return false;
    }
    const value = await readJsonObject(response);
    if (response.status !== 200) {
      throw new BridgeHttpError(response.status, responseMessage(value, "Pending poll failed"));
    }
    // R4d (R3b lane-B A1): the ticket was claimed on a session that is no
    // longer current (Stop, rotation or 401 auto-reconnect while the body was
    // read). It is neither executed nor posted; the companion's session
    // deletion or claim lease owns it. The path from here to executeTool is synchronous.
    if (!this.isCapturedSessionCurrent(session, generation)) return false;
    const ticketId = value.ticketId;
    const requestId = value.requestId;
    const tool = value.tool;
    const argumentsValue = value.arguments;
    const trustedValue = value.trusted;
    if (!isUuid(ticketId)) throw new Error("Companion ticketId must be a UUID");
    // R4f (R4-TX8b): a ticket already dispatched on this session is never
    // re-executed or re-posted; the companion's claim lease or timeout owns it.
    if (this.isDispatchedTicket(session, ticketId)) return false;
    if (!isUuid(requestId)) throw new Error("Companion requestId must be a UUID");
    if (typeof tool !== "string" || !tool || tool.length > 256) {
      throw new Error("Companion tool name is invalid");
    }
    if (!isRecord(argumentsValue)) {
      if (tool === INTERNAL_PLAN_TOOL) {
        await this.completeInvalidPlanTicket(session, ticketId);
        return true;
      }
      throw new Error("Companion arguments must be an object");
    }

    if (tool !== INTERNAL_PLAN_TOOL) {
      // Primitive tickets: unchanged two-argument call and result envelope.
      // R4g-2a (r2d-1 A3): a primitive ticket that carries trusted plan
      // metadata is answered with the fixed primitive error; it is never
      // executed and never recorded as dispatched.
      if (trustedValue !== undefined) {
        const refused: CompletedResult = {
          result: { ok: false, category: "Execution", error: "Office tool execution failed" },
          owner: session,
        };
        this.completedResults.set(ticketId, refused);
        await this.flushCompletedResult(session, ticketId, refused);
        return true;
      }
      let result: unknown;
      try {
        this.activeToolCalls += 1;
        this.recordDispatchedTicket(session, ticketId);
        const raw = await this.deps.executeTool(tool, argumentsValue);
        try {
          result = JSON.parse(raw);
        } catch {
          result = raw;
        }
      } catch {
        result = {
          ok: false,
          category: "Execution",
          error: "Office tool execution failed",
        };
      } finally {
        this.activeToolCalls -= 1;
      }
      const completed: CompletedResult = { result, owner: session };
      this.completedResults.set(ticketId, completed);
      await this.flushCompletedResult(session, ticketId, completed);
      return true;
    }

    // Trusted plan tickets. A malformed ticket is answered with an explicit
    // non-executing disposition instead of waiting out the claim lease.
    let trusted: BridgeTrustedPlanEnvelope;
    try {
      const parsed = parseTrustedPlanEnvelope(trustedValue, session.workbookPlanFingerprint);
      if (argumentsValue.version !== 1) {
        throw new Error("Companion plan schema version is invalid");
      }
      trusted = handedPlanEnvelope(parsed, argumentsValue.outcome);
    } catch {
      await this.completeInvalidPlanTicket(session, ticketId);
      return true;
    }

    const acknowledgement = this.createExecuteStartAcknowledgement(session, generation, ticketId);
    let result: unknown;
    try {
      // R4g-2a (R4j B-A1): counted and recorded inside the try, as on the primitive path.
      this.activeToolCalls += 1;
      this.recordDispatchedTicket(session, ticketId);
      const raw = await this.deps.executeTool(tool, argumentsValue, trusted, acknowledgement.control);
      try {
        result = JSON.parse(raw);
      } catch {
        result = raw;
      }
    } catch {
      result = acknowledgement.acknowledged()
        ? { ok: false, action: "workbook_plan", category: "Execution", state: "uncertain", error: PLAN_UNCONFIRMED_MESSAGE }
        : { ok: false, action: "workbook_plan", category: "Execution", reason: "rejected", error: PLAN_NOT_EXECUTED_MESSAGE };
    } finally {
      acknowledgement.close();
      this.activeToolCalls -= 1;
    }

    const completed: CompletedResult = { result, executed: acknowledgement.acknowledged(), owner: session };
    this.completedResults.set(ticketId, completed);
    await this.flushCompletedResult(session, ticketId, completed);
    return true;
  }

  private async completeInvalidPlanTicket(session: PrivateSession, ticketId: string): Promise<void> {
    const completed: CompletedResult = {
      result: { ok: false, action: "workbook_plan", reason: "invalid", error: INVALID_PLAN_TICKET_MESSAGE },
      executed: false,
      owner: session,
    };
    this.completedResults.set(ticketId, completed);
    await this.flushCompletedResult(session, ticketId, completed);
  }

  /**
   * Builds the one-shot execute-start acknowledgement for one trusted plan
   * ticket. It is bound to the exact PrivateSession object and lifecycle
   * generation captured at poll time, so Stop or a 401 auto-reconnect (which
   * replaces the session without bumping the generation) rejects it locally.
   */
  private createExecuteStartAcknowledgement(
    session: PrivateSession,
    generation: number,
    ticketId: string,
  ): { readonly control: BridgePlanExecutionControl; readonly acknowledged: () => boolean; readonly close: () => void } {
    let used = false;
    let closed = false;
    let acknowledged = false;
    const control: BridgePlanExecutionControl = Object.freeze({
      acknowledgeExecuteStart: (signal?: AbortSignal): Promise<void> => {
        if (used || closed) return Promise.reject(new Error(ACK_REJECTED_MESSAGE));
        used = true;
        return this.postExecuteStart(session, generation, ticketId, signal).then(() => {
          if (closed || !this.isCapturedSessionCurrent(session, generation)) {
            throw new Error(ACK_REJECTED_MESSAGE);
          }
          acknowledged = true;
        });
      },
    });
    return {
      control,
      acknowledged: () => acknowledged,
      close: () => { closed = true; },
    };
  }

  private isCapturedSessionCurrent(session: PrivateSession, generation: number): boolean {
    return this.session === session && this.lifecycleGeneration === generation;
  }

  private isDispatchedTicket(session: PrivateSession, ticketId: string): boolean {
    return this.dispatchedTickets.get(session)?.includes(ticketId) === true;
  }

  /** Synchronous, immediately before executeTool; bounded per session (FIFO). */
  private recordDispatchedTicket(session: PrivateSession, ticketId: string): void {
    let dispatched = this.dispatchedTickets.get(session);
    if (dispatched === undefined) {
      dispatched = [];
      this.dispatchedTickets.set(session, dispatched);
    }
    dispatched.push(ticketId);
    while (dispatched.length > MAX_DISPATCHED_TICKETS_PER_SESSION) dispatched.shift();
  }

  private postExecuteStart(
    session: PrivateSession,
    generation: number,
    ticketId: string,
    signal?: AbortSignal,
  ): Promise<void> {
    if (!this.isCapturedSessionCurrent(session, generation) || signal?.aborted === true) {
      return Promise.reject(new Error(ACK_REJECTED_MESSAGE));
    }
    return new Promise<void>((resolve, reject) => {
      const controller = new AbortController();
      let settled = false;
      const finish = (ok: boolean): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        signal?.removeEventListener("abort", onAbort);
        if (!ok) controller.abort();
        if (ok) resolve(); else reject(new Error(ACK_REJECTED_MESSAGE));
      };
      const onAbort = (): void => finish(false);
      const timer = setTimeout(() => finish(false), EXECUTE_START_ACK_TIMEOUT_MS);
      signal?.addEventListener("abort", onAbort, { once: true });
      let request: Promise<Response>;
      try {
        request = this.fetchImpl(this.url("execute-start"), {
          method: "POST",
          headers: this.sessionHeaders(session, true),
          body: JSON.stringify({ ticketId }),
          signal: controller.signal,
        });
      } catch {
        finish(false);
        return;
      }
      request.then(
        (response) => {
          // Content-free: the companion's response text never reaches the pane.
          void response.text().catch(() => undefined);
          finish(response.status === 200);
        },
        () => finish(false),
      );
    });
  }

  private async flushCompletedResult(
    session: PrivateSession,
    ticketId: string,
    completed: CompletedResult,
  ): Promise<void> {
    completed.attempts = (completed.attempts ?? 0) + 1;
    try {
      await this.submitResult(session, ticketId, completed.oversize === true ? this.oversizeResult(completed) : completed);
    } catch (error) {
      if (!(error instanceof BridgeHttpError) || ![404, 409].includes(error.status)) {
        // Bounded re-post: the original error still rejects this poll, but the
        // entry is dropped after MAX_RESULT_POST_ATTEMPTS so later tickets are
        // served. `executed` never changes and nothing is re-executed; the
        // companion's lease sweep owns the ticket's final state.
        if ((error instanceof BridgeHttpError && error.status === 413) ||
            (!(error instanceof BridgeHttpError) &&
              utf8ByteLength(resultBodyText(ticketId, completed.oversize === true ? this.oversizeResult(completed) : completed)) > MAX_RESULT_BODY_BYTES)) {
          completed.oversize = true;
        }
        if (completed.attempts >= MAX_RESULT_POST_ATTEMPTS) this.completedResults.delete(ticketId);
        throw error;
      }
    }
    this.completedResults.delete(ticketId);
  }

  /**
   * Fixed, content-free substitute for a result the companion refused as too
   * large (413). Same `executed` value; reuses the bridge's existing literals.
   */
  private oversizeResult(completed: CompletedResult): CompletedResult {
    if (completed.executed === undefined) {
      return { result: { ok: false, category: "Execution", error: OVERSIZE_PRIMITIVE_RESULT_MESSAGE } };
    }
    return {
      result: completed.executed
        ? { ok: false, action: "workbook_plan", category: "Execution", state: "uncertain", error: PLAN_UNCONFIRMED_MESSAGE }
        : { ok: false, action: "workbook_plan", category: "Execution", reason: "rejected", error: PLAN_NOT_EXECUTED_MESSAGE },
      executed: completed.executed,
    };
  }

  private async submitResult(
    session: PrivateSession,
    ticketId: string,
    completed: CompletedResult,
  ): Promise<void> {
    const response = await this.fetchImpl(this.url("result"), {
      method: "POST",
      headers: this.sessionHeaders(session, true),
      body: resultBodyText(ticketId, completed),
    });
    if (response.status !== 200) {
      if (response.status === 413) {
        const value = await readJsonObject(response).catch(() => ({}));
        throw new BridgeHttpError(413, responseMessage(value, "Result submission failed"));
      }
      const errorValue = await readJsonObject(response);
      throw new BridgeHttpError(
        response.status,
        responseMessage(errorValue, "Result submission failed"),
      );
    }
  }

  async stop(): Promise<void> {
    this.lifecycleGeneration += 1;
    this.running = false;
    const connecting = this.connecting;
    if (connecting) {
      try {
        await connecting;
      } catch {
        // Failed registration has no server session to unregister.
      }
    }
    if (this.heartbeatTimer !== null) {
      this.clearIntervalImpl(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
    if (this.pollTimer !== null) {
      this.clearIntervalImpl(this.pollTimer);
      this.pollTimer = null;
    }
    const session = this.session;
    this.session = null;
    this.completedResults.clear();
    this.emitState();
    if (!session) return;
    try {
      await this.fetchImpl(this.url("session"), {
        method: "DELETE",
        headers: this.sessionHeaders(session, false),
        keepalive: true,
      });
    } catch {
      // Heartbeat TTL remains the authoritative cleanup path.
    }
  }

  private async handleBackgroundError(error: unknown, session: PrivateSession | null, generation: number): Promise<void> {
    if (session !== null && !this.isCapturedSessionCurrent(session, generation)) return;
    const message = error instanceof Error ? error.message : "Excel bridge unavailable";
    this.deps.onState?.({ connected: false, sessionId: this.sessionId, documentId: this.documentId, lastError: message });
    if (error instanceof BridgeHttpError && error.status === 401 && this.running) {
      this.session = null;
      this.completedResults.clear();
      try {
        await this.connect();
      } catch (reconnectError) {
        const reconnectMessage = reconnectError instanceof Error
          ? reconnectError.message
          : "Excel bridge reconnect failed";
        this.deps.onState?.({ connected: false, sessionId: null, documentId: null, lastError: reconnectMessage });
      }
    }
  }

  chatHeaders(): Record<string, string> | null {
    return this.session ? this.sessionHeaders(this.session, false) : null;
  }

  private sessionHeaders(session: PrivateSession, includeJson: boolean): Record<string, string> {
    return {
      ...(includeJson ? { "Content-Type": "application/json" } : {}),
      "X-Hermes-Session-Key": session.sessionKey,
      "X-Hermes-Session-Id": session.sessionId,
      "X-Hermes-Pane-Id": this.deps.paneId,
    };
  }

  private requireSession(): PrivateSession {
    if (!this.session) throw new Error("Excel bridge session is not connected");
    return this.session;
  }

  private url(route: string): string {
    return `${this.baseUrl}/${this.deps.hostPrefix}/${route}`;
  }

  private publicSession(session: PrivateSession): BridgeSession {
    return {
      sessionId: session.sessionId,
      documentId: session.documentId,
      protocolVersion: session.protocolVersion,
      expiresInSeconds: session.expiresInSeconds,
    };
  }

  private emitState(): void {
    this.deps.onState?.(this.publicState);
  }
}

/**
 * R4g-2a: exactly the four v2 keys; a fresh object in the fixed key order.
 * R4g-2b (R4g-2a lane-B A2): turnOperationId and turnKey are canonical
 * lowercase hyphenated UUIDs (no braces, urn:, uppercase, missing dashes or nil).
 */
function parseTrustedPlanEnvelope(
  value: unknown,
  negotiatedFingerprint: string | null,
): BridgeTrustedPlanEnvelopeV2 {
  if (!isRecord(value)
      || negotiatedFingerprint !== WORKBOOK_PLAN_FINGERPRINT
      || Object.keys(value).length !== 4
      || typeof value.turnOperationId !== "string"
      || !isCanonicalUuid(value.turnOperationId)
      || value.workbookPlanFingerprint !== WORKBOOK_PLAN_FINGERPRINT
      || (value.phase !== "inspect" && value.phase !== "execute")
      || typeof value.turnKey !== "string"
      || !isCanonicalUuid(value.turnKey)
      || value.turnKey === value.turnOperationId) {
    throw new Error("Companion returned an invalid trusted plan envelope");
  }
  return {
    turnOperationId: value.turnOperationId,
    workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
    phase: value.phase,
    turnKey: value.turnKey,
  };
}

/**
 * R4g-2a (R4g-1 lane-B B-A1; OMH correction 2026-10-07 item 2): the phase must
 * agree with the plan outcome (inspect iff outcome "inspect"), else the ticket
 * is invalid. While the route's SERVER_INSPECTION_SUPPORTED is false no
 * phase-bearing envelope reaches executeTool: both phases are handed over as
 * the legacy two-key envelope (the base-tested path; the route refuses a
 * legacy inspect plan before any read, exactly as before R4g-2a).
 */
function handedPlanEnvelope(parsed: BridgeTrustedPlanEnvelopeV2, outcome: unknown): BridgeTrustedPlanEnvelope {
  if ((parsed.phase === "inspect") !== (outcome === "inspect")) {
    throw new Error("Companion returned an invalid trusted plan envelope");
  }
  if ((SERVER_INSPECTION_SUPPORTED as boolean) === true) return parsed;
  return { turnOperationId: parsed.turnOperationId, workbookPlanFingerprint: parsed.workbookPlanFingerprint };
}

/** The exact JSON text POSTed to /excel/result for `completed`. */
function resultBodyText(ticketId: string, completed: CompletedResult): string {
  const body = completed.executed === undefined
    ? { ticketId, result: completed.result }
    : { ticketId, executed: completed.executed, result: completed.result };
  return JSON.stringify(body);
}

function utf8ByteLength(text: string): number {
  return new TextEncoder().encode(text).length;
}

async function readJsonObject(response: Response): Promise<Record<string, unknown>> {
  const declaredLength = response.headers?.get?.("content-length");
  if (declaredLength !== null && declaredLength !== undefined) {
    const normalizedLength = declaredLength.trim();
    if (!/^\d+$/.test(normalizedLength)) throw new Error("Companion response is too large");
    const declaredBytes = Number(normalizedLength);
    if (!Number.isSafeInteger(declaredBytes) || declaredBytes > MAX_RESPONSE_BYTES) {
      throw new Error("Companion response is too large");
    }
  }

  const text = await readBoundedResponseText(response);
  if (!text) return {};
  const value: unknown = JSON.parse(text);
  if (!isRecord(value)) throw new Error("Companion response must be a JSON object");
  return value;
}

async function readBoundedResponseText(response: Response): Promise<string> {
  if (!response.body) {
    const text = await response.text();
    if (utf8ByteLength(text) > MAX_RESPONSE_BYTES) {
      throw new Error("Companion response is too large");
    }
    return text;
  }

  const reader = response.body.getReader();
  // Retain one bounded buffer, not all input chunks plus an assembly copy.
  const bytes = new Uint8Array(MAX_RESPONSE_BYTES);
  let totalBytes = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      if (value.byteLength > MAX_RESPONSE_BYTES - totalBytes) {
        // Cancellation is cleanup only: neither a pending promise nor a failure
        // may delay or replace the established oversize rejection.
        try {
          void reader.cancel().catch(() => undefined);
        } catch {
          // Response doubles may throw synchronously; oversize still wins.
        }
        throw new Error("Companion response is too large");
      }
      bytes.set(value, totalBytes);
      totalBytes += value.byteLength;
    }
  } finally {
    reader.releaseLock();
  }

  return new TextDecoder().decode(bytes.subarray(0, totalBytes));
}

function responseMessage(value: Record<string, unknown>, fallback: string): string {
  return typeof value.error === "string" ? value.error : fallback;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isUuid(value: unknown): value is string {
  return typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}

/** R4g-2b: the canonical spelling only (lowercase; versions 1-5 and the RFC variant exclude nil). */
function isCanonicalUuid(value: unknown): value is string {
  return typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(value);
}
