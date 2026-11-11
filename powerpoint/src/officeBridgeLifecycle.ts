import { BridgePublicState, BridgeSession } from "./officeBridge";

export interface ManagedOfficeBridge {
  readonly sessionId: string | null;
  start(): Promise<BridgeSession>;
  stop(): Promise<void>;
  setStateHandler?(handler: (state: BridgePublicState) => void): void;
}

export interface VisibilityRegistration {
  hiddenValue: unknown;
  register(handler: (message: { visibilityMode: unknown }) => void): Promise<() => void | Promise<void>>;
}

export class OfficeBridgeLifecycleController {
  private bridge: ManagedOfficeBridge | null = null;
  private startPromise: Promise<BridgeSession> | null = null;
  private stopPromise: Promise<void> | null = null;
  private removeVisibilityHandler: (() => void | Promise<void>) | null = null;
  private disposed = false;
  private lifecycleGeneration = 0;
  private visibilityRegistrationGeneration = 0;
  private bridgeStateActive = false;
  private restartWhenStartSettles = false;
  private startWaitingForStop = false;

  constructor(
    private readonly createBridge: (onState: (state: BridgePublicState) => void) => ManagedOfficeBridge,
    private readonly onState: (state: BridgePublicState) => void,
  ) {}

  get sessionId(): string | null {
    if (!this.bridgeStateActive) return null;
    return this.bridge?.sessionId ?? null;
  }

  start(): Promise<BridgeSession> {
    if (this.disposed) {
      return Promise.reject(new Error("Bridge lifecycle is disposed"));
    }
    if (this.startPromise) {
      if (!this.startWaitingForStop && !this.bridgeStateActive) this.restartWhenStartSettles = true;
      return this.startPromise;
    }

    const generation = this.lifecycleGeneration;
    const stop = this.stopPromise;
    let bridge = this.bridge;
    const begin = (): Promise<BridgeSession> => {
      try {
        if (this.disposed) throw new Error("Bridge lifecycle is disposed");
        if (generation !== this.lifecycleGeneration) {
          throw new Error("Bridge lifecycle start was invalidated");
        }
        if (!this.bridge) {
          this.createManagedBridge(generation);
        }
        bridge = this.bridge;
        if (!bridge) throw new Error("Bridge unavailable");
        this.bridgeStateActive = true;
        bridge.setStateHandler?.(this.makeBridgeStateHandler(bridge, generation));
        return Promise.resolve(bridge.start());
      } catch (error) {
        return Promise.reject(error);
      }
    };

    this.startWaitingForStop = stop !== null;
    const underlying = stop
      ? stop.then(
        () => {
          this.startWaitingForStop = false;
          return begin();
        },
        (error: unknown) => {
          this.startWaitingForStop = false;
          return Promise.reject(error);
        },
      )
      : begin();

    let pending: Promise<BridgeSession>;
    pending = Promise.resolve(underlying).then(
      (session) => {
        if (bridge && this.isCurrentBridgeRun(bridge, generation)) {
          this.onState({ connected: true, sessionId: session.sessionId });
        }
        return session;
      },
      (error: unknown) => {
        if (this.isCurrentStartAttempt(bridge, generation)) {
          this.onState({
            connected: false,
            sessionId: null,
            lastError: lifecycleErrorMessage(error),
          });
        }
        throw error;
      },
    );
    this.startPromise = pending;
    void pending.then(
      () => this.finishStart(pending),
      () => this.finishStart(pending),
    );
    return pending;
  }

  stop(): Promise<void> {
    if (this.stopPromise) return this.stopPromise;
    if (!this.bridgeStateActive) return Promise.resolve();
    this.lifecycleGeneration += 1;
    this.bridgeStateActive = false;
    this.onState({ connected: false, sessionId: null });
    const bridge = this.bridge;
    let underlying: Promise<void>;
    try {
      underlying = bridge ? bridge.stop() : Promise.resolve();
    } catch (error) {
      underlying = Promise.reject(error);
    }
    const pending = underlying;
    this.stopPromise = pending;
    void pending.then(
      () => this.finishStop(pending),
      () => this.finishStop(pending),
    );
    return pending;
  }

  register(target: Pick<Window, "addEventListener">, visibility?: VisibilityRegistration): void {
    if (this.disposed) return;
    target.addEventListener("pagehide", () => {
      void this.dispose().catch(() => {
        // Page teardown is best-effort; heartbeat TTL remains authoritative.
      });
    }, { once: true });
    if (!visibility) return;

    const registrationGeneration = ++this.visibilityRegistrationGeneration;
    void visibility.register((message) => {
      if (this.disposed || registrationGeneration !== this.visibilityRegistrationGeneration) return;
      if (message.visibilityMode === visibility.hiddenValue) {
        void this.stop().catch(() => {
          // Stop is best-effort; session TTL cleans up an unreachable companion.
        });
      } else {
        if (this.startPromise) this.restartWhenStartSettles = true;
        void this.start().catch(() => {
          // start() already publishes the disconnected error state.
        });
      }
    }).then((removeHandler) => {
      if (
        this.disposed ||
        registrationGeneration !== this.visibilityRegistrationGeneration
      ) {
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
    if (this.disposed) return;
    this.disposed = true;
    this.visibilityRegistrationGeneration += 1;
    let stopError: unknown;
    try {
      await this.stop();
    } catch (error) {
      stopError = error;
    }
    const removeHandler = this.removeVisibilityHandler;
    this.removeVisibilityHandler = null;
    if (removeHandler) {
      await removeHandler();
    }
    if (stopError) throw stopError;
  }

  private createManagedBridge(runGeneration: number): void {
    let created: ManagedOfficeBridge | null = null;
    const bridgeState = this.makeBridgeStateHandlerForCreated(
      () => created,
      runGeneration,
    );
    created = this.createBridge(bridgeState);
    this.bridge = created;
  }

  private makeBridgeStateHandler(
    bridge: ManagedOfficeBridge,
    generation: number,
  ): (state: BridgePublicState) => void {
    return (state) => {
      if (!this.isCurrentBridgeRun(bridge, generation)) return;
      this.onState(state);
    };
  }

  private makeBridgeStateHandlerForCreated(
    getBridge: () => ManagedOfficeBridge | null,
    generation: number,
  ): (state: BridgePublicState) => void {
    return (state) => {
      const bridge = getBridge();
      if (
        this.disposed ||
        !this.bridgeStateActive ||
        !bridge ||
        this.bridge !== bridge ||
        this.lifecycleGeneration !== generation
      ) {
        return;
      }
      this.onState(state);
    };
  }

  private isCurrentBridgeRun(
    bridge: ManagedOfficeBridge,
    generation: number,
  ): boolean {
    return !this.disposed &&
      this.bridgeStateActive &&
      this.bridge === bridge &&
      this.lifecycleGeneration === generation;
  }

  private isCurrentStartAttempt(
    bridge: ManagedOfficeBridge | null,
    generation: number,
  ): boolean {
    return !this.disposed &&
      generation === this.lifecycleGeneration &&
      (bridge === null ? this.bridge === null : this.bridge === bridge);
  }

  private finishStart(pending: Promise<BridgeSession>): void {
    if (this.startPromise !== pending) return;
    this.startWaitingForStop = false;
    this.startPromise = null;
    if (this.restartWhenStartSettles && !this.disposed && !this.bridgeStateActive) {
      this.restartWhenStartSettles = false;
      void this.start().catch(() => {
        // The visible transition owns the retry; start() publishes its state.
      });
    } else {
      this.restartWhenStartSettles = false;
    }
  }

  private finishStop(pending: Promise<void>): void {
    if (this.stopPromise === pending) this.stopPromise = null;
  }
}

function lifecycleErrorMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : "Companion unavailable";
  if (
    /session.?key|document.?id|pane.?id|ticket.?id|arguments|authorization/i.test(message) ||
    /[A-Za-z0-9+/=_-]{48,}/.test(message)
  ) {
    return "Companion unavailable";
  }
  return message || "Companion unavailable";
}
