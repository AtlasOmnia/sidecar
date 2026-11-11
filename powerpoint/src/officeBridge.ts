import { ApprovalRequiredError, GRANT_HEADER } from "./deskConnect";
const DEFAULT_BASE_URL = "";
const PROTOCOL_VERSION = "1.0";
const HEARTBEAT_MS = 15_000;
const POLL_MS = 250;
const RECONNECT_DELAY_MS = 0;
const REGISTRATION_ATTEMPTS = 5;
const REGISTRATION_TIMEOUT_MS = 10_000;
const REGISTRATION_RETRY_DELAYS = [500, 1_000, 2_000, 4_000] as const;
const RESULT_ATTEMPTS = 5;
const RESULT_TIMEOUT_MS = 10_000;
const RESULT_RETRY_DELAYS = [500, 1_000, 2_000, 4_000] as const;
const MAX_RESULT_BODY_BYTES = 2 * 1024 * 1024;
const MAX_RESPONSE_ENVELOPE_BYTES = 64 * 1024;
const MAX_RESPONSE_BYTES = MAX_RESULT_BODY_BYTES + MAX_RESPONSE_ENVELOPE_BYTES;
const MAX_CACHED_RESULTS = 8;
const MAX_CACHED_RESULT_BYTES = 4 * 1024 * 1024;

interface CachedResult {
  requestId: string;
  ticketId: string;
  originSessionId: string;
  originSessionGeneration: number;
  documentId: string;
  sessionId: string;
  sessionGeneration: number;
  tool: string;
  argumentsFingerprint: string;
  result: unknown;
  serializedBody: string;
  byteLength: number;
  resultAttempts: number;
  retryNotBefore: number | null;
}

interface CachedResultIdentity {
  ticketId: string;
  requestId: string;
  sessionId: string;
  sessionGeneration: number;
  documentId: string;
  tool: string;
  argumentsValue: Record<string, unknown>;
}

type ResultCapacityReservationState = "reserved" | "consumed" | "released";

interface ResultCapacityReservation {
  state: ResultCapacityReservationState;
}

export interface BridgePublicState {
  connected: boolean;
  sessionId: string | null;
  lastError?: string;
}

export interface BridgeSession {
  sessionId: string;
  protocolVersion: string;
  expiresInSeconds: number;
}

interface PrivateSession extends BridgeSession {
  readonly sessionKey: string;
  readonly paneId: string;
  readonly documentId: string;
  readonly sessionGeneration: number;
}

interface RegistrationAttemptOptions {
  readonly documentId: string;
  readonly sessionGeneration: number;
  readonly signal: AbortSignal;
}

const documentExecutionTails = new Map<string, Promise<void>>();

export interface OfficeBridgeDependencies {
  hostPrefix: string;
  documentId: () => Promise<string>;
  paneId: string;
  executeTool: (name: string, argumentsValue: Record<string, unknown>) => Promise<string>;
  fetchImpl?: typeof fetch;
  baseUrl?: string;
  onState?: (state: BridgePublicState) => void;
  grant?: () => string | null;
  onApprovalRequired?: () => void;
  setIntervalImpl?: typeof setInterval;
  clearIntervalImpl?: typeof clearInterval;
  setTimeoutImpl?: typeof setTimeout;
  clearTimeoutImpl?: typeof clearTimeout;
  randomImpl?: () => number;
  nowImpl?: () => number;
}

class BridgeHttpError extends Error {
  constructor(
    public readonly status: number,
    message: string,
    public readonly responseBody?: Record<string, unknown>,
  ) {
    super(message);
    this.name = "BridgeHttpError";
  }
}

class BridgeTimeoutError extends Error {
  constructor() {
    super("Companion registration attempt timed out");
    this.name = "BridgeTimeoutError";
  }
}

class BridgeResultTimeoutError extends Error {
  constructor() {
    super("Companion result submission attempt timed out");
    this.name = "BridgeResultTimeoutError";
  }
}

class BridgeProtocolError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BridgeProtocolError";
  }
}

class BridgeCancelledError extends Error {
  constructor() {
    super("PowerPoint bridge operation was cancelled");
    this.name = "BridgeCancelledError";
  }
}

class BridgeCapacityError extends Error {
  constructor() {
    super("Companion result retry cache capacity exceeded");
    this.name = "BridgeCapacityError";
  }
}

class BridgeValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BridgeValidationError";
  }
}

export class OfficeCompanionBridge {
  private readonly fetchImpl: typeof fetch;
  private readonly baseUrl: string;
  private readonly setIntervalImpl: typeof setInterval;
  private readonly clearIntervalImpl: typeof clearInterval;
  private readonly setTimeoutImpl: typeof setTimeout;
  private readonly clearTimeoutImpl: typeof clearTimeout;
  private readonly randomImpl: () => number;
  private readonly nowImpl: () => number;
  private stateHandler: ((state: BridgePublicState) => void) | undefined;
  private session: PrivateSession | null = null;
  private registrationFlight: Promise<BridgeSession> | null = null;
  private registrationController: AbortController | null = null;
  private sessionAbortController: AbortController | null = null;
  private readonly completedResults = new Map<string, CachedResult>();
  private completedResultBytes = 0;
  private reservedResultSlots = 0;
  private reservedResultBytes = 0;
  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;
  private pollTimer: ReturnType<typeof setInterval> | null = null;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private retryTimer: ReturnType<typeof setTimeout> | null = null;
  private resultRetryTimer: ReturnType<typeof setTimeout> | null = null;
  private cancelRetryWait: (() => void) | null = null;
  private cancelResultRetryWait: (() => void) | null = null;
  private pollFlight: Promise<boolean> | null = null;
  private heartbeatFlight: Promise<void> | null = null;
  private running = false;
  private lifecycleGeneration = 0;
  private sessionGeneration = 0;
  private readonly unregisteredLeaseKeys = new Set<string>();

  constructor(private readonly deps: OfficeBridgeDependencies) {
    const fetchImpl = deps.fetchImpl ?? globalThis.fetch;
    this.fetchImpl = fetchImpl.bind(globalThis);
    this.baseUrl = (deps.baseUrl ?? DEFAULT_BASE_URL).replace(/\/+$/, "");
    const setIntervalImpl = deps.setIntervalImpl ?? globalThis.setInterval;
    const clearIntervalImpl = deps.clearIntervalImpl ?? globalThis.clearInterval;
    const setTimeoutImpl = deps.setTimeoutImpl ?? globalThis.setTimeout;
    const clearTimeoutImpl = deps.clearTimeoutImpl ?? globalThis.clearTimeout;
    this.setIntervalImpl = setIntervalImpl.bind(globalThis);
    this.clearIntervalImpl = clearIntervalImpl.bind(globalThis);
    this.setTimeoutImpl = setTimeoutImpl.bind(globalThis);
    this.clearTimeoutImpl = clearTimeoutImpl.bind(globalThis);
    this.randomImpl = deps.randomImpl ?? (() => Math.random());
    this.nowImpl = deps.nowImpl ?? (() => Date.now());
    this.stateHandler = deps.onState;
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

  get publicState(): BridgePublicState {
    return { connected: this.session !== null, sessionId: this.sessionId };
  }

  setStateHandler(handler: (state: BridgePublicState) => void): void {
    this.stateHandler = handler;
  }

  connect(): Promise<BridgeSession> {
    if (this.session) {
      return Promise.resolve(this.publicSession(this.session));
    }
    if (this.registrationFlight) {
      return this.registrationFlight;
    }

    this.clearReconnectTimer();
    const sessionGeneration = this.reserveSessionGeneration();
    const controller = new AbortController();
    this.registrationController = controller;
    const connecting = this.connectNewSession(sessionGeneration, controller);
    this.registrationFlight = connecting;
    void connecting.then(
      () => {
        if (this.registrationFlight === connecting) {
          this.registrationFlight = null;
          this.registrationController = null;
        }
      },
      () => {
        if (this.registrationFlight === connecting) {
          this.registrationFlight = null;
          this.registrationController = null;
        }
      },
    );
    return connecting;
  }

  private async connectNewSession(
    sessionGeneration: number,
    controller: AbortController,
  ): Promise<BridgeSession> {
    const documentId = await awaitWithSignal(
      Promise.resolve(this.deps.documentId()),
      controller.signal,
    );
    if (!this.isCurrentRegistration(sessionGeneration, controller)) {
      throw new BridgeCancelledError();
    }
    if (!documentId || documentId.length > 256) {
      throw new BridgeValidationError("documentId must be a bounded non-empty string");
    }

    for (let attempt = 0; attempt < REGISTRATION_ATTEMPTS; attempt += 1) {
      if (!this.isCurrentRegistration(sessionGeneration, controller)) {
        throw new BridgeCancelledError();
      }
      if (attempt > 0) {
        const baseDelay = REGISTRATION_RETRY_DELAYS[attempt - 1];
        const jitter = Math.floor(clampRandom(this.randomImpl()) * 250);
        await this.waitForRetry(baseDelay + jitter, controller.signal);
        if (!this.isCurrentRegistration(sessionGeneration, controller)) {
          throw new BridgeCancelledError();
        }
      }

      try {
        const lease = await this.registerAttempt({
          documentId,
          sessionGeneration,
          signal: controller.signal,
        });
        if (!this.isCurrentRegistration(sessionGeneration, controller)) {
          void this.unregisterLease(lease);
          throw new BridgeCancelledError();
        }
        this.installSession(lease);
        return this.publicSession(lease);
      } catch (error) {
        if (!this.isCurrentRegistration(sessionGeneration, controller)) {
          throw new BridgeCancelledError();
        }
        if (attempt + 1 < REGISTRATION_ATTEMPTS && isRetryableRegistrationError(error)) {
          continue;
        }
        throw error;
      }
    }

    throw new Error("Companion registration failed");
  }

  private async registerAttempt(options: RegistrationAttemptOptions): Promise<PrivateSession> {
    const attemptController = new AbortController();
    const onCampaignAbort = () => attemptController.abort();
    if (options.signal.aborted) {
      attemptController.abort();
    } else {
      options.signal.addEventListener("abort", onCampaignAbort, { once: true });
    }
    const operation = this.performRegistrationAttempt({
      ...options,
      signal: attemptController.signal,
    });
    try {
      return await this.awaitRegistrationAttempt(
        operation,
        options.signal,
        () => attemptController.abort(),
      );
    } catch (error) {
      if (error instanceof BridgeCancelledError || error instanceof BridgeTimeoutError) {
        let cleaned = false;
        void operation.then(
          (lease) => {
            if (cleaned) return;
            cleaned = true;
            void this.unregisterLease(lease);
          },
          () => {
            // The late attempt has no validated lease to clean up.
          },
        );
      }
      throw error;
    } finally {
      options.signal.removeEventListener("abort", onCampaignAbort);
    }
  }

  private async performRegistrationAttempt(
    options: RegistrationAttemptOptions,
  ): Promise<PrivateSession> {
    const grant = this.deps.grant?.() ?? null;
    let response: Response;
    try {
      response = await this.fetchImpl(this.url("session"), {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(grant ? { [GRANT_HEADER]: grant } : {}) },
        body: JSON.stringify({
          protocolVersion: PROTOCOL_VERSION,
          documentId: options.documentId,
          paneId: this.deps.paneId,
        }),
        credentials: "omit",
        redirect: "error",
        signal: options.signal,
      });
    } catch (error) {
      throw error;
    }

    let value: Record<string, unknown>;
    try {
      value = await readJsonObject(response);
    } catch (error) {
      throw new BridgeValidationError(
        error instanceof Error ? error.message : "Companion response is invalid",
      );
    }
    if (response.status === 403 && value.approval === "required") {
      this.deps.onApprovalRequired?.();
      throw new ApprovalRequiredError();
    }
    if (response.status !== 201) {
      throw new BridgeHttpError(response.status, "Session registration failed");
    }

    const sessionId = value.sessionId;
    const sessionKey = value.sessionKey;
    const protocolVersion = value.protocolVersion;
    const expiresInSeconds = value.expiresInSeconds;
    if (!isUuid(sessionId) || typeof sessionKey !== "string" || sessionKey.length < 32) {
      throw new BridgeValidationError("Companion returned an invalid session identity");
    }
    if (
      protocolVersion !== PROTOCOL_VERSION ||
      typeof expiresInSeconds !== "number" ||
      !Number.isFinite(expiresInSeconds)
    ) {
      throw new BridgeValidationError("Companion returned an incompatible protocol response");
    }

    return Object.freeze({
      sessionId,
      sessionKey,
      protocolVersion,
      expiresInSeconds,
      paneId: this.deps.paneId,
      documentId: options.documentId,
      sessionGeneration: options.sessionGeneration,
    });
  }

  private awaitRegistrationAttempt<T>(
    operation: Promise<T>,
    signal: AbortSignal,
    onTimeout: () => void,
  ): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      let settled = false;
      let timer: ReturnType<typeof setTimeout> | null = null;
      const cleanup = () => {
        signal.removeEventListener("abort", onAbort);
        if (timer !== null) {
          this.clearTimeoutImpl(timer);
          timer = null;
        }
      };
      const finish = (callback: () => void) => {
        if (settled) return;
        settled = true;
        cleanup();
        callback();
      };
      const onAbort = () => finish(() => reject(new BridgeCancelledError()));

      if (signal.aborted) {
        void Promise.resolve(operation).catch(() => {
          // Observe a fetch that ignores cancellation.
        });
        reject(new BridgeCancelledError());
        return;
      }
      signal.addEventListener("abort", onAbort, { once: true });
      timer = this.setTimeoutImpl(
        () => finish(() => {
          onTimeout();
          reject(new BridgeTimeoutError());
        }),
        REGISTRATION_TIMEOUT_MS,
      );
      Promise.resolve(operation).then(
        (value) => finish(() => resolve(value)),
        (error) => finish(() => reject(error)),
      );
    });
  }

  private waitForRetry(delayMs: number, signal: AbortSignal): Promise<void> {
    if (signal.aborted) return Promise.reject(new BridgeCancelledError());
    if (this.retryTimer !== null) {
      return Promise.reject(new Error("Registration retry timer already exists"));
    }
    return new Promise<void>((resolve, reject) => {
      let settled = false;
      let timer: ReturnType<typeof setTimeout> | null = null;
      const cleanup = () => {
        signal.removeEventListener("abort", onAbort);
        if (this.cancelRetryWait === cancel) this.cancelRetryWait = null;
      };
      const finish = (callback: () => void) => {
        if (settled) return;
        settled = true;
        if (timer !== null) {
          this.clearTimeoutImpl(timer);
          timer = null;
        }
        if (this.retryTimer !== null) this.retryTimer = null;
        cleanup();
        callback();
      };
      const onAbort = () => finish(() => reject(new BridgeCancelledError()));
      const cancel = () => finish(() => reject(new BridgeCancelledError()));
      this.cancelRetryWait = cancel;
      signal.addEventListener("abort", onAbort, { once: true });
      timer = this.setTimeoutImpl(() => finish(resolve), delayMs);
      this.retryTimer = timer;
    });
  }

  async start(): Promise<BridgeSession> {
    const lifecycleGeneration = this.lifecycleGeneration;
    const session = await this.connect();
    const lease = this.session;
    if (
      lifecycleGeneration !== this.lifecycleGeneration ||
      !lease ||
      lease.sessionId !== session.sessionId ||
      !this.isCurrentLease(lease)
    ) {
      throw new BridgeCancelledError();
    }
    if (this.running) {
      return session;
    }
    this.running = true;
    this.ensureTimers();
    return session;
  }

  heartbeatOnce(): Promise<void> {
    const lease = this.requireSession();
    const flight = this.beginHeartbeat(lease);
    void flight.catch((error) => this.handleBackgroundError(error, lease, true));
    return flight;
  }

  private beginHeartbeat(lease: PrivateSession): Promise<void> {
    if (this.heartbeatFlight) return this.heartbeatFlight;
    const flight = this.heartbeatForLease(lease).catch((error: unknown) => {
      if (error instanceof BridgeCancelledError) return;
      throw error;
    });
    this.heartbeatFlight = flight;
    void flight.then(
      () => {
        if (this.heartbeatFlight === flight) this.heartbeatFlight = null;
      },
      () => {
        if (this.heartbeatFlight === flight) this.heartbeatFlight = null;
      },
    );
    return flight;
  }

  private async heartbeatForLease(lease: PrivateSession): Promise<void> {
    const signal = this.signalForLease(lease);
    if (!signal) return;
    const response = await awaitWithSignal(
      Promise.resolve(this.fetchImpl(this.url("heartbeat"), {
        method: "POST",
        headers: this.sessionHeaders(lease, true),
        body: "{}",
        credentials: "omit",
        redirect: "error",
        signal,
      })),
      signal,
    );
    if (!this.isCurrentLease(lease)) return;
    const value = await awaitWithSignal(readJsonObject(response), signal);
    if (!this.isCurrentLease(lease)) return;
    if (response.status !== 200) {
      throw new BridgeHttpError(response.status, "Heartbeat failed");
    }
    void value;
  }

  pollOnce(): Promise<boolean> {
    if (this.pollFlight) return this.pollFlight;
    const lease = this.requireSession();
    const flight = this.pollForLease(lease).catch((error: unknown) => {
      if (error instanceof BridgeCancelledError) return false;
      throw error;
    });
    this.pollFlight = flight;
    void flight.then(
      () => {
        if (this.pollFlight === flight) this.pollFlight = null;
      },
      () => {
        if (this.pollFlight === flight) this.pollFlight = null;
      },
    );
    void flight.catch((error) => this.handleBackgroundError(error, lease, true));
    return flight;
  }

  private async pollForLease(lease: PrivateSession): Promise<boolean> {
    const signal = this.signalForLease(lease);
    if (!signal || !this.isCurrentLease(lease)) return false;

    const response = await awaitWithSignal(
      Promise.resolve(this.fetchImpl(this.url("pending"), {
        method: "GET",
        headers: this.sessionHeaders(lease, false),
        credentials: "omit",
        redirect: "error",
        signal,
      })),
      signal,
    );
    if (!this.isCurrentLease(lease)) return false;
    if (response.status === 204) {
      const cached = this.findCachedResultForLease(lease);
      return cached ? this.flushCompletedResult(lease, cached) : false;
    }

    const value = await awaitWithSignal(readJsonObject(response), signal);
    if (!this.isCurrentLease(lease)) return false;
    if (response.status !== 200) {
      throw new BridgeHttpError(response.status, "Pending poll failed", value);
    }

    const ticketId = value.ticketId;
    const requestId = value.requestId;
    const tool = value.tool;
    const argumentsValue = value.arguments;
    if (!isUuid(ticketId)) throw new BridgeValidationError("Companion ticketId must be a UUID");
    if (!isUuid(requestId)) throw new BridgeValidationError("Companion requestId must be a UUID");
    if (typeof tool !== "string" || !tool || tool.length > 256) {
      throw new BridgeValidationError("Companion tool name is invalid");
    }
    if (!isRecord(argumentsValue)) {
      throw new BridgeValidationError("Companion arguments must be an object");
    }
    if (!this.isCurrentLease(lease)) return false;

    const cached = this.completedResults.get(requestId);
    if (cached) {
      if (!this.matchesCachedRequest(cached, lease, tool, argumentsValue)) {
        throw new BridgeValidationError("Companion request conflicts with a retained result");
      }
      const updated = this.reattachCachedResult(cached, lease, ticketId);
      return this.flushCompletedResult(lease, updated);
    }

    const cachedResult = await this.executeAndCacheResult(
      lease,
      signal,
      ticketId,
      requestId,
      tool,
      argumentsValue,
    );
    if (!cachedResult || !this.isCurrentLease(lease)) return false;
    return this.flushCompletedResult(lease, cachedResult);
  }

  private async executeAndCacheResult(
    lease: PrivateSession,
    signal: AbortSignal,
    ticketId: string,
    requestId: string,
    tool: string,
    argumentsValue: Record<string, unknown>,
  ): Promise<CachedResult | null> {
    let cachedAfterWait: CachedResult | null = null;
    let reservation: ResultCapacityReservation | null = null;
    let result: unknown;
    try {
      reservation = this.reserveResultCapacity();
      result = await this.withDocumentExecution(lease.documentId, signal, async () => {
        const existing = this.completedResults.get(requestId);
        if (existing) {
          if (!this.matchesCachedRequest(existing, lease, tool, argumentsValue)) {
            throw new BridgeValidationError("Companion request conflicts with a retained result");
          }
          cachedAfterWait = existing;
          return undefined;
        }
        return this.runToolExecution(lease, signal, tool, argumentsValue);
      });

      if (cachedAfterWait !== null) {
        this.releaseResultCapacity(reservation);
        reservation = null;
        return this.reattachCachedResult(cachedAfterWait, lease, ticketId);
      }

      const cached = this.cacheCompletedResult({
        ticketId,
        requestId,
        sessionId: lease.sessionId,
        sessionGeneration: lease.sessionGeneration,
        documentId: lease.documentId,
        tool,
        argumentsValue,
      }, result, reservation);
      reservation = null;
      return cached;
    } catch (error) {
      if (reservation !== null) this.releaseResultCapacity(reservation);
      if (error instanceof BridgeCancelledError && !this.isCurrentLease(lease)) return null;
      throw error;
    }
  }

  private async runToolExecution(
    lease: PrivateSession,
    signal: AbortSignal,
    tool: string,
    argumentsValue: Record<string, unknown>,
  ): Promise<unknown> {
    if (!this.isCurrentLease(lease)) throw new BridgeCancelledError();

    let execution: Promise<string>;
    try {
      execution = Promise.resolve(this.deps.executeTool(tool, argumentsValue));
    } catch (error) {
      return executionFailureResult(error);
    }

    try {
      const raw = await awaitWithSignal(execution, signal);
      return parseToolResult(raw);
    } catch (error) {
      if (!this.isCurrentLease(lease)) {
        const settled = await settleOperation(execution);
        return settled.ok ? parseToolResult(settled.value) : executionFailureResult(settled.error);
      }
      return executionFailureResult(error);
    }
  }

  private findCachedResultForLease(lease: PrivateSession): CachedResult | null {
    for (const cached of this.completedResults.values()) {
      if (
        cached.documentId === lease.documentId &&
        cached.sessionId === lease.sessionId &&
        cached.sessionGeneration === lease.sessionGeneration
      ) {
        return cached;
      }
    }
    return null;
  }

  private matchesCachedRequest(
    cached: CachedResult,
    lease: PrivateSession,
    tool: string,
    argumentsValue: Record<string, unknown>,
  ): boolean {
    return cached.documentId === lease.documentId &&
      cached.tool === tool &&
      cached.argumentsFingerprint === canonicalJson(argumentsValue);
  }

  private reattachCachedResult(
    cached: CachedResult,
    lease: PrivateSession,
    ticketId: string,
  ): CachedResult {
    const serialized = serializeResultBody(ticketId, cached.result);
    const updated: CachedResult = {
      ...cached,
      ticketId,
      sessionId: lease.sessionId,
      sessionGeneration: lease.sessionGeneration,
      result: serialized.result,
      serializedBody: serialized.serializedBody,
      byteLength: serialized.byteLength,
    };
    this.replaceCachedResult(updated);
    return updated;
  }

  private async withDocumentExecution<T>(
    documentId: string,
    signal: AbortSignal,
    operation: () => Promise<T>,
  ): Promise<T> {
    const previous = documentExecutionTails.get(documentId) ?? Promise.resolve();
    let release!: () => void;
    const current = new Promise<void>((resolve) => { release = resolve; });
    const chain = previous.then(() => current);
    documentExecutionTails.set(documentId, chain);
    void chain.then(() => {
      if (documentExecutionTails.get(documentId) === chain) {
        documentExecutionTails.delete(documentId);
      }
    });

    try {
      await awaitWithSignal(previous, signal);
      return await operation();
    } finally {
      release();
    }
  }

  private reserveResultCapacity(): ResultCapacityReservation {
    if (
      this.completedResults.size + this.reservedResultSlots + 1 > MAX_CACHED_RESULTS ||
      this.completedResultBytes + this.reservedResultBytes + MAX_RESULT_BODY_BYTES > MAX_CACHED_RESULT_BYTES
    ) {
      throw new BridgeCapacityError();
    }
    this.reservedResultSlots += 1;
    this.reservedResultBytes += MAX_RESULT_BODY_BYTES;
    return { state: "reserved" };
  }

  private releaseResultCapacity(reservation: ResultCapacityReservation): void {
    if (reservation.state !== "reserved") return;
    this.reservedResultSlots -= 1;
    this.reservedResultBytes -= MAX_RESULT_BODY_BYTES;
    reservation.state = "released";
  }

  private consumeResultCapacity(reservation: ResultCapacityReservation): void {
    if (reservation.state !== "reserved") {
      throw new Error("Companion result capacity reservation is not active");
    }
    this.reservedResultSlots -= 1;
    this.reservedResultBytes -= MAX_RESULT_BODY_BYTES;
    reservation.state = "consumed";
  }

  private cacheCompletedResult(
    identityOrTicketId: CachedResultIdentity | string,
    result: unknown,
    reservation?: ResultCapacityReservation,
  ): CachedResult {
    const identity: CachedResultIdentity = typeof identityOrTicketId === "string"
      ? {
          ticketId: identityOrTicketId,
          requestId: identityOrTicketId,
          sessionId: "",
          sessionGeneration: 0,
          documentId: "",
          tool: "",
          argumentsValue: {},
        }
      : identityOrTicketId;
    const existing = this.completedResults.get(identity.requestId);
    if (existing && !reservation) return existing;

    const serialized = serializeResultBody(identity.ticketId, result);
    const cached: CachedResult = {
      requestId: identity.requestId,
      ticketId: identity.ticketId,
      originSessionId: identity.sessionId,
      originSessionGeneration: identity.sessionGeneration,
      documentId: identity.documentId,
      sessionId: identity.sessionId,
      sessionGeneration: identity.sessionGeneration,
      tool: identity.tool,
      argumentsFingerprint: canonicalJson(identity.argumentsValue),
      result: serialized.result,
      serializedBody: serialized.serializedBody,
      byteLength: serialized.byteLength,
      resultAttempts: 0,
      retryNotBefore: null,
    };
    this.replaceCachedResult(cached, reservation);
    return cached;
  }

  private replaceCachedResult(
    cached: CachedResult,
    reservation?: ResultCapacityReservation,
  ): void {
    const existing = this.completedResults.get(cached.requestId);
    const nextCount = this.completedResults.size + (existing ? 0 : 1);
    const nextBytes = this.completedResultBytes - (existing?.byteLength ?? 0) + cached.byteLength;
    if (nextCount > MAX_CACHED_RESULTS || nextBytes > MAX_CACHED_RESULT_BYTES) {
      throw new BridgeCapacityError();
    }

    if (reservation) this.consumeResultCapacity(reservation);
    this.completedResults.set(cached.requestId, cached);
    this.completedResultBytes = nextBytes;
  }

  private deleteCompletedResult(requestId: string): void {
    const cached = this.completedResults.get(requestId);
    if (!cached) return;
    this.completedResultBytes -= cached.byteLength;
    this.completedResults.delete(requestId);
  }

  private async flushCompletedResult(
    lease: PrivateSession,
    cachedResult: CachedResult,
  ): Promise<boolean> {
    if (!this.isCurrentLease(lease)) return false;
    if (cachedResult.resultAttempts >= RESULT_ATTEMPTS) return false;

    while (cachedResult.resultAttempts < RESULT_ATTEMPTS) {
      if (!this.isCurrentLease(lease)) return false;
      if (cachedResult.retryNotBefore !== null) {
        const remaining = Math.max(0, cachedResult.retryNotBefore - this.nowImpl());
        if (remaining > 0) {
          try {
            await this.waitForResultRetry(remaining, this.signalForLease(lease)!);
          } catch (error) {
            if (!this.isCurrentLease(lease)) return false;
            throw error;
          }
        }
        cachedResult.retryNotBefore = null;
      }
      if (!this.isCurrentLease(lease)) return false;

      cachedResult.resultAttempts += 1;
      try {
        await this.submitResult(lease, cachedResult.serializedBody, cachedResult.ticketId);
      } catch (error) {
        if (!this.isCurrentLease(lease)) return false;
        if (isTerminalResultRejection(error)) {
          if (this.completedResults.get(cachedResult.requestId) === cachedResult) {
            this.deleteCompletedResult(cachedResult.requestId);
            return true;
          }
          return false;
        }
        if (!isRetryableResultError(error) || cachedResult.resultAttempts >= RESULT_ATTEMPTS) {
          throw error;
        }
        const baseDelay = RESULT_RETRY_DELAYS[cachedResult.resultAttempts - 1];
        const jitter = Math.floor(clampRandom(this.randomImpl()) * 250);
        cachedResult.retryNotBefore = this.nowImpl() + baseDelay + jitter;
        try {
          await this.waitForResultRetry(baseDelay + jitter, this.signalForLease(lease)!);
        } catch (waitError) {
          if (!this.isCurrentLease(lease)) return false;
          throw waitError;
        }
        if (!this.isCurrentLease(lease)) return false;
        cachedResult.retryNotBefore = null;
        continue;
      }

      if (!this.isCurrentLease(lease)) return false;
      if (this.completedResults.get(cachedResult.requestId) !== cachedResult) return false;
      this.deleteCompletedResult(cachedResult.requestId);
      return true;
    }

    return false;
  }

  private waitForResultRetry(delayMs: number, signal: AbortSignal): Promise<void> {
    if (signal.aborted) return Promise.reject(new BridgeCancelledError());
    if (this.resultRetryTimer !== null) {
      return Promise.reject(new Error("Result retry timer already exists"));
    }
    return new Promise<void>((resolve, reject) => {
      let settled = false;
      let timer: ReturnType<typeof setTimeout> | null = null;
      const cleanup = () => {
        signal.removeEventListener("abort", onAbort);
        if (this.cancelResultRetryWait === cancel) this.cancelResultRetryWait = null;
      };
      const finish = (callback: () => void) => {
        if (settled) return;
        settled = true;
        if (timer !== null) {
          this.clearTimeoutImpl(timer);
          timer = null;
        }
        if (this.resultRetryTimer !== null) this.resultRetryTimer = null;
        cleanup();
        callback();
      };
      const onAbort = () => finish(() => reject(new BridgeCancelledError()));
      const cancel = () => finish(() => reject(new BridgeCancelledError()));
      this.cancelResultRetryWait = cancel;
      signal.addEventListener("abort", onAbort, { once: true });
      timer = this.setTimeoutImpl(() => finish(resolve), delayMs);
      this.resultRetryTimer = timer;
    });
  }

  private async submitResult(
    lease: PrivateSession,
    serializedBody: string,
    ticketId: string,
  ): Promise<void> {
    const leaseSignal = this.signalForLease(lease);
    if (!leaseSignal) return;
    const attemptController = new AbortController();
    const operation = (async (): Promise<void> => {
      const response = await this.fetchImpl(this.url("result"), {
        method: "POST",
        headers: this.sessionHeaders(lease, true),
        body: serializedBody,
        credentials: "omit",
        redirect: "error",
        signal: attemptController.signal,
      });
      if (!this.isCurrentLease(lease)) return;
      const value = await awaitWithSignal(
        readJsonObject(response),
        attemptController.signal,
      );
      if (!this.isCurrentLease(lease)) return;
      if (response.status !== 200) {
        throw new BridgeHttpError(response.status, "Result submission failed", value);
      }
      if (value.ok !== true || value.ticketId !== ticketId) {
        throw new BridgeProtocolError("Companion result acknowledgement was invalid");
      }
    })();

    await this.awaitResultAttempt(
      operation,
      leaseSignal,
      () => attemptController.abort(),
    );
  }

  private awaitResultAttempt<T>(
    operation: Promise<T>,
    signal: AbortSignal,
    onTimeout: () => void,
  ): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      let settled = false;
      let timer: ReturnType<typeof setTimeout> | null = null;
      const cleanup = () => {
        signal.removeEventListener("abort", onAbort);
        if (timer !== null) {
          this.clearTimeoutImpl(timer);
          timer = null;
        }
      };
      const finish = (callback: () => void) => {
        if (settled) return;
        settled = true;
        cleanup();
        callback();
      };
      const onAbort = () => finish(() => {
        onTimeout();
        reject(new BridgeCancelledError());
      });
      if (signal.aborted) {
        void Promise.resolve(operation).catch(() => undefined);
        onTimeout();
        reject(new BridgeCancelledError());
        return;
      }
      signal.addEventListener("abort", onAbort, { once: true });
      timer = this.setTimeoutImpl(
        () => finish(() => {
          onTimeout();
          reject(new BridgeResultTimeoutError());
        }),
        RESULT_TIMEOUT_MS,
      );
      Promise.resolve(operation).then(
        (value) => finish(() => resolve(value)),
        (error) => finish(() => reject(error)),
      );
    });
  }

  async stop(): Promise<void> {
    this.lifecycleGeneration += 1;
    this.running = false;
    this.clearReconnectTimer();
    this.cancelRetryWait?.();
    this.cancelRetryWait = null;
    this.cancelResultRetryWait?.();
    this.cancelResultRetryWait = null;
    const registrationController = this.registrationController;
    registrationController?.abort();
    this.registrationController = null;
    this.registrationFlight = null;
    this.sessionAbortController?.abort();
    this.clearTimers();
    this.pollFlight = null;
    this.heartbeatFlight = null;

    const session = this.session;
    this.session = null;
    this.sessionGeneration += 1;
    this.sessionAbortController = null;
    this.emitState();
    if (session) {
      await this.unregisterLease(session);
    }
  }

  async dispose(): Promise<void> {
    await this.stop();
  }

  private ensureTimers(): void {
    if (this.heartbeatTimer === null) {
      this.heartbeatTimer = this.setIntervalImpl(() => {
        const lease = this.session;
        if (!this.running || !lease || !this.isCurrentLease(lease)) return;
        const flight = this.beginHeartbeat(lease);
        void flight.catch((error) => this.handleBackgroundError(error, lease, true));
      }, HEARTBEAT_MS);
    }
    if (this.pollTimer === null) {
      this.pollTimer = this.setIntervalImpl(() => {
        const lease = this.session;
        if (!this.running || !lease || !this.isCurrentLease(lease)) return;
        const flight = this.pollOnce();
        void flight.catch((error) => this.handleBackgroundError(error, lease, true));
      }, POLL_MS);
    }
  }

  private clearTimers(): void {
    if (this.heartbeatTimer !== null) {
      this.clearIntervalImpl(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
    if (this.pollTimer !== null) {
      this.clearIntervalImpl(this.pollTimer);
      this.pollTimer = null;
    }
  }

  private handleBackgroundError(
    error: unknown,
    lease: PrivateSession,
    reconnectAllowed: boolean,
  ): void {
    if (!this.isCurrentLease(lease) || error instanceof BridgeCancelledError) return;
    if (reconnectAllowed && this.running && isRetryableSessionError(error)) {
      this.invalidateSession(lease);
      this.scheduleReconnect();
      return;
    }
    this.stateHandler?.({
      connected: false,
      sessionId: this.session?.sessionId ?? null,
      lastError: backgroundErrorMessage(error),
    });
  }

  private invalidateSession(lease: PrivateSession): void {
    if (!this.isCurrentLease(lease)) return;
    this.sessionAbortController?.abort();
    this.sessionAbortController = null;
    this.session = null;
    this.sessionGeneration += 1;
    this.pollFlight = null;
    this.heartbeatFlight = null;
    this.emitState();
  }

  private scheduleReconnect(): void {
    if (!this.running || this.session || this.registrationFlight || this.reconnectTimer !== null) {
      return;
    }
    const lifecycleGeneration = this.lifecycleGeneration;
    this.reconnectTimer = this.setTimeoutImpl(() => {
      this.reconnectTimer = null;
      if (
        !this.running ||
        lifecycleGeneration !== this.lifecycleGeneration ||
        this.session ||
        this.registrationFlight
      ) {
        return;
      }
      const reconnect = this.connect();
      void reconnect.catch((error) => {
        if (lifecycleGeneration !== this.lifecycleGeneration || !this.running) return;
        this.stateHandler?.({
          connected: false,
          sessionId: null,
          lastError: backgroundErrorMessage(error),
        });
      });
    }, RECONNECT_DELAY_MS);
  }

  private clearReconnectTimer(): void {
    if (this.reconnectTimer === null) return;
    this.clearTimeoutImpl(this.reconnectTimer);
    this.reconnectTimer = null;
  }

  private reserveSessionGeneration(): number {
    this.sessionGeneration += 1;
    return this.sessionGeneration;
  }

  private isCurrentRegistration(
    sessionGeneration: number,
    controller: AbortController,
  ): boolean {
    return this.sessionGeneration === sessionGeneration &&
      this.registrationController === controller &&
      !controller.signal.aborted;
  }

  private isCurrentLease(lease: PrivateSession): boolean {
    return this.session === lease &&
      this.session?.sessionId === lease.sessionId &&
      this.sessionGeneration === lease.sessionGeneration;
  }

  private installSession(lease: PrivateSession): void {
    if (!this.isCurrentRegistration(lease.sessionGeneration, this.registrationController!)) {
      throw new BridgeCancelledError();
    }
    this.sessionAbortController = new AbortController();
    this.session = lease;
    this.emitState();
  }

  private signalForLease(lease: PrivateSession): AbortSignal | null {
    if (!this.isCurrentLease(lease) || !this.sessionAbortController) return null;
    return this.sessionAbortController.signal;
  }

  private requireSession(): PrivateSession {
    if (!this.session) throw new Error("PowerPoint bridge session is not connected");
    return this.session;
  }

  chatHeaders(): Record<string, string> | null {
    return this.session ? this.sessionHeaders(this.session, false) : null;
  }

  private sessionHeaders(session: PrivateSession, includeJson: boolean): Record<string, string> {
    return {
      ...(includeJson ? { "Content-Type": "application/json" } : {}),
      "X-Hermes-Session-Key": session.sessionKey,
      "X-Hermes-Session-Id": session.sessionId,
      "X-Hermes-Pane-Id": session.paneId,
    };
  }

  private async unregisterLease(session: PrivateSession): Promise<void> {
    const key = `${session.sessionId}\u0000${session.sessionKey}`;
    if (this.unregisteredLeaseKeys.has(key)) return;
    this.unregisteredLeaseKeys.add(key);
    try {
      await this.fetchImpl(this.url("session"), {
        method: "DELETE",
        headers: this.sessionHeaders(session, false),
        keepalive: true,
        credentials: "omit",
        redirect: "error",
      });
    } catch {
      // Heartbeat TTL remains the authoritative cleanup path.
    }
  }

  private url(route: string): string {
    return `${this.baseUrl}/${this.deps.hostPrefix}/${route}`;
  }

  private publicSession(session: PrivateSession): BridgeSession {
    return {
      sessionId: session.sessionId,
      protocolVersion: session.protocolVersion,
      expiresInSeconds: session.expiresInSeconds,
    };
  }

  private emitState(): void {
    this.stateHandler?.(this.publicState);
  }
}

type SettledOperation<T> =
  | { ok: true; value: T }
  | { ok: false; error: unknown };

async function settleOperation<T>(operation: PromiseLike<T>): Promise<SettledOperation<T>> {
  return Promise.resolve(operation).then(
    (value) => ({ ok: true, value }),
    (error) => ({ ok: false, error }),
  );
}

function parseToolResult(raw: string): unknown {
  try {
    return JSON.parse(raw);
  } catch {
    return raw;
  }
}

function executionFailureResult(error: unknown): Record<string, unknown> {
  const message = error instanceof Error ? error.message : "Office tool execution failed";
  return {
    ok: false,
    category: "Execution",
    error: sanitizeExecutionError(message),
  };
}

function sanitizeExecutionError(message: string): string {
  if (
    /session.?key|document.?id|pane.?id|ticket.?id|authorization/i.test(message) ||
    /[A-Za-z0-9+/=_-]{48,}/.test(message)
  ) {
    return "Office tool execution failed";
  }
  return message.slice(0, 512) || "Office tool execution failed";
}

function serializeResultBody(
  ticketId: string,
  result: unknown,
): { result: unknown; serializedBody: string; byteLength: number } {
  try {
    const serializedBody = JSON.stringify({ ticketId, result });
    if (typeof serializedBody === "string") {
      const byteLength = utf8ByteLength(serializedBody);
      if (byteLength <= MAX_RESULT_BODY_BYTES) {
        return { result, serializedBody, byteLength };
      }
    }
  } catch {
    // Replace a post-execution serialization failure with a bounded result.
  }

  const safeResult = {
    ok: false,
    category: "Result",
    error: "Companion result is too large",
  };
  const serializedBody = JSON.stringify({ ticketId, result: safeResult });
  return {
    result: safeResult,
    serializedBody,
    byteLength: utf8ByteLength(serializedBody),
  };
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalize(value)) ?? "undefined";
}

function canonicalize(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((item) => canonicalize(item));
  if (!isRecord(value)) return value;
  const normalized: Record<string, unknown> = {};
  for (const key of Object.keys(value).sort()) {
    normalized[key] = canonicalize(value[key]);
  }
  return normalized;
}

async function awaitWithSignal<T>(operation: PromiseLike<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) {
    void Promise.resolve(operation).catch(() => {
      // Observe a non-compliant operation after cancellation.
    });
    throw new BridgeCancelledError();
  }
  return new Promise<T>((resolve, reject) => {
    let settled = false;
    const cleanup = () => signal.removeEventListener("abort", onAbort);
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      cleanup();
      callback();
    };
    const onAbort = () => finish(() => reject(new BridgeCancelledError()));
    signal.addEventListener("abort", onAbort, { once: true });
    Promise.resolve(operation).then(
      (value) => finish(() => resolve(value)),
      (error) => finish(() => reject(error)),
    );
  });
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
  const chunks: Uint8Array[] = [];
  let totalBytes = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      totalBytes += value.byteLength;
      if (totalBytes > MAX_RESPONSE_BYTES) {
        await reader.cancel();
        throw new Error("Companion response is too large");
      }
      chunks.push(value);
    }
  } finally {
    reader.releaseLock();
  }

  const bytes = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return new TextDecoder().decode(bytes);
}

function utf8ByteLength(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function isTerminalResultRejection(error: unknown): boolean {
  if (!(error instanceof BridgeHttpError) || !error.responseBody) return false;
  if (error.responseBody.ok !== false) return false;
  const category = error.responseBody.category;
  return (error.status === 400 && category === "Validation") ||
    (error.status === 409 && category === "Conflict") ||
    (error.status === 499 && category === "Cancellation");
}

function isRetryableResultError(error: unknown): boolean {
  if (error instanceof BridgeCancelledError || error instanceof BridgeProtocolError) return false;
  if (isTerminalResultRejection(error)) return false;
  return true;
}

function isRetryableRegistrationError(error: unknown): boolean {
  if (error instanceof ApprovalRequiredError || error instanceof BridgeCancelledError || error instanceof BridgeValidationError) return false;
  if (error instanceof BridgeHttpError) {
    return error.status === 401 ||
      error.status === 404 ||
      error.status === 408 ||
      error.status === 425 ||
      error.status === 429 ||
      error.status >= 500;
  }
  return error instanceof BridgeTimeoutError || error instanceof TypeError || error instanceof Error;
}

function isRetryableSessionError(error: unknown): boolean {
  if (error instanceof BridgeCancelledError || error instanceof BridgeCapacityError || error instanceof BridgeValidationError) return false;
  if (error instanceof BridgeHttpError) {
    return error.status === 401 ||
      error.status === 404 ||
      error.status === 408 ||
      error.status === 425 ||
      error.status === 429 ||
      error.status >= 500;
  }
  return error instanceof BridgeTimeoutError || error instanceof TypeError || error instanceof Error;
}

function backgroundErrorMessage(error: unknown): string {
  if (error instanceof BridgeTimeoutError) return "Companion request timed out";
  if (error instanceof BridgeHttpError) return `Companion request failed (${error.status})`;
  return "PowerPoint bridge unavailable";
}

function clampRandom(value: number): number {
  if (!Number.isFinite(value) || value <= 0) return 0;
  if (value >= 1) return Number.MIN_VALUE + (1 - Number.EPSILON);
  return value;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isUuid(value: unknown): value is string {
  return typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
