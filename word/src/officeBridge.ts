import { ApprovalRequiredError, GRANT_HEADER } from "./deskConnect";
const DEFAULT_BASE_URL = "";
const PROTOCOL_VERSION = "1.0";
const HEARTBEAT_MS = 15_000;
const POLL_MS = 250;
const MAX_RESULT_BODY_BYTES = 10 * 1024 * 1024;
const MAX_RESPONSE_ENVELOPE_BYTES = 64 * 1024;
const MAX_RESPONSE_BYTES = MAX_RESULT_BODY_BYTES + MAX_RESPONSE_ENVELOPE_BYTES;
const MAX_CACHED_RESULTS = 4;
const MAX_CACHED_RESULT_BYTES = 12 * 1024 * 1024;

interface CachedResult {
  result: unknown;
  serializedBody: string;
  byteLength: number;
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
  sessionKey: string;
}

export interface OfficeBridgeDependencies {
  hostPrefix: string;
  documentId: () => Promise<string>;
  paneId: string;
  executeTool: (name: string, argumentsValue: Record<string, unknown>) => Promise<string>;
  fetchImpl?: typeof fetch;
  baseUrl?: string;
  onState?: (state: BridgePublicState) => void;
  /** Sidecar grant for this sidebar (sent only when registering a session). */
  grant?: () => string | null;
  /** Called when Hermes no longer (or not yet) approves this sidebar. */
  onApprovalRequired?: () => void;
  setIntervalImpl?: typeof setInterval;
  clearIntervalImpl?: typeof clearInterval;
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
  private readonly completedResults = new Map<string, CachedResult>();
  private completedResultBytes = 0;
  private heartbeatTimer: ReturnType<typeof setInterval> | null = null;
  private pollTimer: ReturnType<typeof setInterval> | null = null;
  private polling = false;
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

  get publicState(): BridgePublicState {
    return { connected: this.session !== null, sessionId: this.sessionId };
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
    this.session = { sessionId, sessionKey, protocolVersion, expiresInSeconds };
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
      void this.heartbeatOnce().catch((error) => this.handleBackgroundError(error));
    }, HEARTBEAT_MS);
    this.pollTimer = this.setIntervalImpl(() => {
      if (this.polling) return;
      this.polling = true;
      void this.pollOnce()
        .catch((error) => this.handleBackgroundError(error))
        .finally(() => { this.polling = false; });
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
    const cached = this.completedResults.entries().next();
    if (!cached.done) {
      const [ticketId, cachedResult] = cached.value;
      await this.submitCachedResult(session, ticketId, cachedResult);
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
    const ticketId = value.ticketId;
    const requestId = value.requestId;
    const tool = value.tool;
    const argumentsValue = value.arguments;
    if (!isUuid(ticketId)) throw new Error("Companion ticketId must be a UUID");
    if (!isUuid(requestId)) throw new Error("Companion requestId must be a UUID");
    if (typeof tool !== "string" || !tool || tool.length > 256) {
      throw new Error("Companion tool name is invalid");
    }
    if (!isRecord(argumentsValue)) {
      throw new Error("Companion arguments must be an object");
    }

    let result: unknown;
    try {
      const raw = await this.deps.executeTool(tool, argumentsValue);
      try {
        result = JSON.parse(raw);
      } catch {
        result = raw;
      }
    } catch (error) {
      result = {
        ok: false,
        category: "Execution",
        error: error instanceof Error ? error.message : "Office tool execution failed",
      };
    }

    const cachedResult = this.cacheCompletedResult(ticketId, result);
    await this.submitCachedResult(session, ticketId, cachedResult);
    return true;
  }

  private cacheCompletedResult(ticketId: string, result: unknown): CachedResult {
    const serializedBody = JSON.stringify({ ticketId, result });
    const byteLength = utf8ByteLength(serializedBody);
    if (byteLength > MAX_RESULT_BODY_BYTES) {
      throw new Error("Companion result is too large");
    }

    const existing = this.completedResults.get(ticketId);
    const existingBytes = existing?.byteLength ?? 0;
    const nextCount = this.completedResults.size + (existing ? 0 : 1);
    const nextBytes = this.completedResultBytes - existingBytes + byteLength;
    if (nextCount > MAX_CACHED_RESULTS || nextBytes > MAX_CACHED_RESULT_BYTES) {
      throw new Error("Companion result retry cache capacity exceeded");
    }

    const cached = { result, serializedBody, byteLength };
    this.completedResults.set(ticketId, cached);
    this.completedResultBytes = nextBytes;
    return cached;
  }

  private deleteCompletedResult(ticketId: string): void {
    const cached = this.completedResults.get(ticketId);
    if (!cached) return;
    this.completedResultBytes -= cached.byteLength;
    this.completedResults.delete(ticketId);
  }

  private clearCompletedResults(): void {
    this.completedResults.clear();
    this.completedResultBytes = 0;
  }

  private async submitCachedResult(
    session: PrivateSession,
    ticketId: string,
    cachedResult: CachedResult,
  ): Promise<void> {
    try {
      await this.submitResult(session, cachedResult.serializedBody);
      this.deleteCompletedResult(ticketId);
    } catch (error) {
      if (isTerminalResultRejection(error)) {
        this.deleteCompletedResult(ticketId);
      }
      throw error;
    }
  }

  private async submitResult(
    session: PrivateSession,
    serializedBody: string,
  ): Promise<void> {
    const response = await this.fetchImpl(this.url("result"), {
      method: "POST",
      headers: this.sessionHeaders(session, true),
      body: serializedBody,
    });
    if (response.status !== 200) {
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
    this.clearCompletedResults();
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

  private async handleBackgroundError(error: unknown): Promise<void> {
    const message = error instanceof Error ? error.message : "Office bridge unavailable";
    this.deps.onState?.({ connected: false, sessionId: this.sessionId, lastError: message });
    if (error instanceof BridgeHttpError && error.status === 401 && this.running) {
      this.session = null;
      this.clearCompletedResults();
      try {
        await this.connect();
      } catch (reconnectError) {
        const reconnectMessage = reconnectError instanceof Error
          ? reconnectError.message
          : "Office bridge reconnect failed";
        this.deps.onState?.({ connected: false, sessionId: null, lastError: reconnectMessage });
      }
    }
  }

  /** Headers that authenticate this sidebar's chat requests at the bridge (null when not connected). */
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
    if (!this.session) throw new Error("Office bridge session is not connected");
    return this.session;
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
    this.deps.onState?.(this.publicState);
  }
}

async function readJsonObject(response: Response): Promise<Record<string, unknown>> {
  const contentLength = response.headers?.get?.("content-length");
  if (contentLength !== null && contentLength !== undefined) {
    const declaredBytes = Number(contentLength);
    if (Number.isFinite(declaredBytes) && declaredBytes > MAX_RESPONSE_BYTES) {
      throw new Error("Companion response is too large");
    }
  }

  const text = await response.text();
  if (utf8ByteLength(text) > MAX_RESPONSE_BYTES) {
    throw new Error("Companion response is too large");
  }
  if (!text) return {};
  const value: unknown = JSON.parse(text);
  if (!isRecord(value)) throw new Error("Companion response must be a JSON object");
  return value;
}

function utf8ByteLength(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function responseMessage(value: Record<string, unknown>, fallback: string): string {
  return typeof value.error === "string" ? value.error : fallback;
}

function isTerminalResultRejection(error: unknown): boolean {
  return error instanceof BridgeHttpError
    && error.status >= 400
    && error.status < 500
    && ![408, 425, 429].includes(error.status);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isUuid(value: unknown): value is string {
  return typeof value === "string" &&
    /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
