import { BridgePublicState, BridgeSession } from "./officeBridge";

export interface ManagedOfficeBridge {
  readonly sessionId: string | null;
  readonly documentId: string | null;
  start(): Promise<BridgeSession>;
  stop(): Promise<void>;
}

export interface VisibilityRegistration {
  hiddenValue: unknown;
  register(handler: (message: { visibilityMode: unknown }) => void): Promise<() => void | Promise<void>>;
}

export class OfficeBridgeLifecycleController {
  private bridge: ManagedOfficeBridge | null = null;
  private startPromise: Promise<BridgeSession> | null = null;
  private removeVisibilityHandler: (() => void | Promise<void>) | null = null;
  private disposed = false;
  private lifecycleGeneration = 0;

  constructor(
    private readonly createBridge: (onState: (state: BridgePublicState) => void) => ManagedOfficeBridge,
    private readonly onState: (state: BridgePublicState) => void,
  ) {}

  get sessionId(): string | null {
    return this.bridge?.sessionId ?? null;
  }

  get documentId(): string | null {
    return this.bridge?.documentId ?? null;
  }

  async start(): Promise<BridgeSession> {
    if (this.disposed) {
      throw new Error("Bridge lifecycle is disposed");
    }
    if (!this.bridge) {
      this.bridge = this.createBridge(this.onState);
    }
    if (this.startPromise) {
      return this.startPromise;
    }
    const generation = this.lifecycleGeneration;
    const pending = this.bridge.start();
    this.startPromise = pending;
    try {
      const session = await pending;
      if (!this.disposed && generation === this.lifecycleGeneration) {
        this.onState({ connected: true, sessionId: session.sessionId, documentId: session.documentId });
      }
      return session;
    } catch (error) {
      if (!this.disposed && generation === this.lifecycleGeneration) {
        this.onState({
          connected: false,
          sessionId: null,
          documentId: null,
          lastError: error instanceof Error ? error.message : "Companion unavailable",
        });
      }
      throw error;
    } finally {
      if (this.startPromise === pending) {
        this.startPromise = null;
      }
    }
  }

  async stop(): Promise<void> {
    this.lifecycleGeneration += 1;
    this.startPromise = null;
    await this.bridge?.stop();
  }

  async reconnectForDocumentChange(expectedDocumentId: string): Promise<BridgeSession> {
    if (!expectedDocumentId) throw new Error("Expected document identity is unavailable");
    await this.stop();
    const session = await this.start();
    if (session.documentId !== expectedDocumentId || this.documentId !== expectedDocumentId) {
      await this.stop();
      throw new Error("Bridge session document identity does not match the current document");
    }
    return session;
  }

  register(target: Pick<Window, "addEventListener">, visibility?: VisibilityRegistration): void {
    this.disposed = false;
    target.addEventListener("pagehide", () => {
      void this.dispose().catch(() => {
        // Page teardown is best-effort; heartbeat TTL remains authoritative.
      });
    }, { once: true });
    if (!visibility) return;

    void visibility.register((message) => {
      if (message.visibilityMode === visibility.hiddenValue) {
        void this.stop().catch(() => {
          // Stop is best-effort; session TTL cleans up an unreachable companion.
        });
      } else {
        void this.start().catch(() => {
          // start() already publishes the disconnected error state.
        });
      }
    }).then((removeHandler) => {
      if (this.disposed) {
        void Promise.resolve(removeHandler()).catch(() => {
          // Handler cleanup during teardown is best-effort.
        });
      } else {
        this.removeVisibilityHandler = removeHandler;
      }
    }).catch(() => {
      // SharedRuntime visibility events are optional; heartbeat expiry is authoritative.
    });
  }

  async dispose(): Promise<void> {
    this.disposed = true;
    await this.stop();
    const removeHandler = this.removeVisibilityHandler;
    this.removeVisibilityHandler = null;
    if (removeHandler) {
      await removeHandler();
    }
  }
}
