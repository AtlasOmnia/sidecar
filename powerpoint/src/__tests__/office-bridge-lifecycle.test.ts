import { describe, expect, it, vi } from "vitest";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";

const SESSION = { sessionId: "660e8400-e29b-41d4-a716-446655440001", protocolVersion: "1.0", expiresInSeconds: 60 };

function flush(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function eventTargetHarness() {
  const listeners = new Map<string, EventListener>();
  return {
    target: {
      addEventListener(type: string, listener: EventListenerOrEventListenerObject) {
        listeners.set(type, listener as EventListener);
      },
    } as Pick<Window, "addEventListener">,
    fire(type: string) {
      listeners.get(type)?.(new Event(type));
    },
  };
}

describe("OfficeBridgeLifecycleController", () => {
  it("deduplicates concurrent initial starts and publishes connected state", async () => {
    let resolveStart!: (value: typeof SESSION) => void;
    const start = vi.fn(() => new Promise<typeof SESSION>((resolve) => { resolveStart = resolve; }));
    const bridge = { sessionId: SESSION.sessionId, start, stop: vi.fn(async () => {}) };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);

    const first = controller.start();
    const second = controller.start();
    expect(start).toHaveBeenCalledTimes(1);
    resolveStart(SESSION);
    await expect(Promise.all([first, second])).resolves.toEqual([SESSION, SESSION]);
    expect(onState).toHaveBeenLastCalledWith({ connected: true, sessionId: SESSION.sessionId });
  });

  it("suppresses a late connected state after stop invalidates a pending start", async () => {
    let resolveStart!: (value: typeof SESSION) => void;
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(() => new Promise<typeof SESSION>((resolve) => { resolveStart = resolve; })),
      stop: vi.fn(async () => {}),
    };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);

    const pending = controller.start();
    await controller.stop();
    resolveStart(SESSION);
    await pending;

    expect(bridge.stop).toHaveBeenCalledTimes(1);
    expect(onState).not.toHaveBeenCalledWith({ connected: true, sessionId: SESSION.sessionId });
  });

  it("stops when hidden and restarts when visible without creating another bridge", async () => {
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const factory = vi.fn(() => bridge);
    const controller = new OfficeBridgeLifecycleController(factory, vi.fn());
    const harness = eventTargetHarness();
    let visibilityHandler!: (message: { visibilityMode: unknown }) => void;
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async (handler) => { visibilityHandler = handler; return vi.fn(); },
    });

    await controller.start();
    visibilityHandler({ visibilityMode: "hidden" });
    await flush();
    visibilityHandler({ visibilityMode: "visible" });
    await flush();

    expect(bridge.stop).toHaveBeenCalledTimes(1);
    expect(bridge.start).toHaveBeenCalledTimes(2);
    expect(factory).toHaveBeenCalledTimes(1);
  });

  it("pagehide stops the bridge and removes the visibility handler", async () => {
    const removeHandler = vi.fn(async () => {});
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());
    const harness = eventTargetHarness();
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async () => removeHandler,
    });
    await controller.start();
    await flush();

    harness.fire("pagehide");
    await flush();
    expect(bridge.stop).toHaveBeenCalledTimes(1);
    expect(removeHandler).toHaveBeenCalledTimes(1);
  });

  it("reports start failures without exposing credential material", async () => {
    const onState = vi.fn();
    const bridge = {
      sessionId: null,
      start: vi.fn(async () => { throw new Error("companion unavailable"); }),
      stop: vi.fn(async () => {}),
    };
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);

    await expect(controller.start()).rejects.toThrow("companion unavailable");
    expect(onState).toHaveBeenLastCalledWith({
      connected: false,
      sessionId: null,
      lastError: "companion unavailable",
    });
    expect(JSON.stringify(onState.mock.calls)).not.toContain("sessionKey");
  });


  it("suppresses bridge state callbacks captured before stop invalidates the lifecycle", async () => {
    let bridgeState!: (state: { connected: boolean; sessionId: string | null }) => void;
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController((callback) => {
      bridgeState = callback;
      return bridge;
    }, onState);

    await controller.start();
    await controller.stop();
    const stateCountAfterStop = onState.mock.calls.length;
    bridgeState({ connected: true, sessionId: SESSION.sessionId });

    expect(onState.mock.calls.length).toBe(stateCountAfterStop);
  });


  it("keeps startPromise coalesced until the underlying start settles", async () => {
    let resolveStart!: (value: typeof SESSION) => void;
    const start = vi.fn(() => new Promise<typeof SESSION>((resolve) => { resolveStart = resolve; }));
    const bridge = { sessionId: SESSION.sessionId, start, stop: vi.fn(async () => {}) };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());

    const first = controller.start();
    await controller.stop();
    const overlapping = controller.start();

    expect(start).toHaveBeenCalledTimes(1);
    expect(overlapping).toBe(first);
    resolveStart(SESSION);
    await expect(first).resolves.toEqual(SESSION);
  });


  it("waits for a pending stop before starting the bridge", async () => {
    let resolveStop!: () => void;
    const start = vi.fn(async () => SESSION);
    const stop = vi.fn(() => new Promise<void>((resolve) => { resolveStop = resolve; }));
    const bridge = { sessionId: SESSION.sessionId, start, stop };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());

    await controller.start();
    const stopping = controller.stop();
    const restarting = controller.start();
    const repeated = controller.start();

    expect(repeated).toBe(restarting);
    expect(start).toHaveBeenCalledTimes(1);

    resolveStop();
    await stopping;
    await expect(restarting).resolves.toEqual(SESSION);
    expect(start).toHaveBeenCalledTimes(2);
  });

  it("contains a synchronous factory failure in the rejected start promise", async () => {
    const secret = `sessionKey=${"s".repeat(64)}`;
    let factoryCalls = 0;
    const start = vi.fn(async () => SESSION);
    const bridge = { sessionId: SESSION.sessionId, start, stop: vi.fn(async () => {}) };
    const factory = vi.fn(() => {
      factoryCalls += 1;
      if (factoryCalls === 1) throw new Error(`factory failed: ${secret}`);
      return bridge;
    });
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(factory, onState);

    let first!: Promise<typeof SESSION>;
    expect(() => { first = controller.start(); }).not.toThrow();
    await expect(first).rejects.toThrow(secret);
    expect(onState).toHaveBeenLastCalledWith({
      connected: false,
      sessionId: null,
      lastError: "Companion unavailable",
    });
    expect(JSON.stringify(onState.mock.calls)).not.toContain(secret);

    await expect(controller.start()).resolves.toEqual(SESSION);
    expect(factory).toHaveBeenCalledTimes(2);
    expect(start).toHaveBeenCalledTimes(1);
  });

  it("fails a queued start closed when the preceding stop rejects", async () => {
    let rejectStop!: (error: unknown) => void;
    const start = vi.fn(async () => SESSION);
    const stop = vi.fn(() => new Promise<void>((_resolve, reject) => { rejectStop = reject; }));
    const bridge = { sessionId: SESSION.sessionId, start, stop };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);

    await controller.start();
    const stopping = controller.stop();
    const restarting = controller.start();
    const repeated = controller.start();

    expect(repeated).toBe(restarting);
    expect(start).toHaveBeenCalledTimes(1);
    const stopError = new Error("stop failed");
    rejectStop(stopError);

    await expect(stopping).rejects.toBe(stopError);
    await expect(restarting).rejects.toBe(stopError);
    expect(start).toHaveBeenCalledTimes(1);
  });

  it("does not start after disposal while waiting for a pending stop", async () => {
    let resolveStop!: () => void;
    const start = vi.fn(async () => SESSION);
    const stop = vi.fn(() => new Promise<void>((resolve) => { resolveStop = resolve; }));
    const bridge = { sessionId: SESSION.sessionId, start, stop };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());

    await controller.start();
    const stopping = controller.stop();
    const restarting = controller.start();
    const disposing = controller.dispose();

    expect(start).toHaveBeenCalledTimes(1);
    resolveStop();
    await expect(stopping).resolves.toBeUndefined();
    await expect(restarting).rejects.toThrow("disposed");
    await expect(disposing).resolves.toBeUndefined();
    expect(start).toHaveBeenCalledTimes(1);
  });

  it("removes a late visibility registration after pagehide disposal", async () => {
    let resolveRegistration!: (value: () => void | Promise<void>) => void;
    let visibilityHandler!: (message: { visibilityMode: unknown }) => void;
    const removeHandler = vi.fn(async () => {});
    const bridge = {
      sessionId: null,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());
    const harness = eventTargetHarness();
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async (handler) => {
        visibilityHandler = handler;
        return new Promise((resolve) => { resolveRegistration = resolve; });
      },
    });

    harness.fire("pagehide");
    await flush();
    resolveRegistration(removeHandler);
    await flush();
    await flush();

    expect(removeHandler).toHaveBeenCalledTimes(1);
    visibilityHandler({ visibilityMode: "visible" });
    await flush();
    expect(bridge.start).not.toHaveBeenCalled();
  });

  it("suppresses overlapping visible start and pagehide completion", async () => {
    let resolveStart!: (value: typeof SESSION) => void;
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(() => new Promise<typeof SESSION>((resolve) => { resolveStart = resolve; })),
      stop: vi.fn(async () => {}),
    };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);
    const harness = eventTargetHarness();
    let visibilityHandler!: (message: { visibilityMode: unknown }) => void;
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async (handler) => { visibilityHandler = handler; return vi.fn(); },
    });

    const first = controller.start();
    visibilityHandler({ visibilityMode: "visible" });
    harness.fire("pagehide");
    resolveStart(SESSION);
    await expect(first).resolves.toEqual(SESSION);
    await flush();

    expect(bridge.start).toHaveBeenCalledTimes(1);
    expect(bridge.stop).toHaveBeenCalledTimes(1);
    expect(onState).not.toHaveBeenCalledWith({ connected: true, sessionId: SESSION.sessionId });
  });


  it("restarts after a visible reset arrives while a stale start is still pending", async () => {
    let resolveFirstStart!: (value: typeof SESSION) => void;
    let startCount = 0;
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(() => {
        startCount += 1;
        if (startCount === 1) {
          return new Promise<typeof SESSION>((resolve) => { resolveFirstStart = resolve; });
        }
        return Promise.resolve(SESSION);
      }),
      stop: vi.fn(async () => {}),
    };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());
    const harness = eventTargetHarness();
    let visibilityHandler!: (message: { visibilityMode: unknown }) => void;
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async (handler) => { visibilityHandler = handler; return vi.fn(); },
    });

    const first = controller.start();
    await controller.stop();
    await flush();
    visibilityHandler({ visibilityMode: "visible" });
    expect(bridge.start).toHaveBeenCalledTimes(1);
    resolveFirstStart(SESSION);
    await expect(first).resolves.toEqual(SESSION);
    await flush();
    await flush();

    expect(bridge.start).toHaveBeenCalledTimes(2);
  });


  it("removes visibility handling even when bridge stop fails during disposal", async () => {
    const removeHandler = vi.fn(async () => {});
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => { throw new Error("stop failed"); }),
    };
    const controller = new OfficeBridgeLifecycleController(() => bridge, vi.fn());
    const harness = eventTargetHarness();
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async () => removeHandler,
    });
    await controller.start();
    await flush();

    await expect(controller.dispose()).rejects.toThrow("stop failed");
    expect(removeHandler).toHaveBeenCalledTimes(1);
  });

  it("wraps a synchronous bridge start failure in a rejected promise", async () => {
    const onState = vi.fn();
    const bridge = {
      sessionId: null,
      start: vi.fn(() => { throw new Error("synchronous start failed"); }),
      stop: vi.fn(async () => {}),
    };
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);

    let pending!: Promise<typeof SESSION>;
    expect(() => { pending = controller.start(); }).not.toThrow();
    await expect(pending).rejects.toThrow("synchronous start failed");
    expect(onState).toHaveBeenLastCalledWith({
      connected: false,
      sessionId: null,
      lastError: "synchronous start failed",
    });
  });

  it("suppresses a factory callback fired before the bridge is adopted", async () => {
    const onState = vi.fn();
    const factory = vi.fn((bridgeState: (state: { connected: boolean; sessionId: string | null }) => void) => {
      bridgeState({ connected: true, sessionId: "99999999-9999-4999-8999-999999999999" });
      return {
        sessionId: SESSION.sessionId,
        start: vi.fn(async () => SESSION),
        stop: vi.fn(async () => {}),
      };
    });
    const controller = new OfficeBridgeLifecycleController(factory, onState);

    await controller.start();
    expect(onState).not.toHaveBeenCalledWith({
      connected: true,
      sessionId: "99999999-9999-4999-8999-999999999999",
    });
  });

  it("suppresses a pre-stop bridge callback after the same bridge starts a new run and emits disconnected on stop", async () => {
    let bridgeState!: (state: { connected: boolean; sessionId: string | null }) => void;
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController((callback) => {
      bridgeState = callback;
      return bridge;
    }, onState);

    await controller.start();
    const preStopCallback = bridgeState;
    await controller.stop();
    expect(onState).toHaveBeenLastCalledWith({ connected: false, sessionId: null });

    await controller.start();
    const stateCountAfterRestart = onState.mock.calls.length;
    preStopCallback({ connected: true, sessionId: SESSION.sessionId });
    expect(onState.mock.calls.length).toBe(stateCountAfterRestart);
  });

  it("emits one disconnected state and stops once per active run", async () => {
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);

    await controller.start();
    await controller.stop();
    await controller.stop();

    expect(bridge.stop).toHaveBeenCalledTimes(1);
    expect(onState.mock.calls.filter(([state]) =>
      state.connected === false && state.sessionId === null,
    )).toHaveLength(1);

    await controller.start();
    await controller.stop();
    await controller.stop();

    expect(bridge.stop).toHaveBeenCalledTimes(2);
    expect(onState.mock.calls.filter(([state]) =>
      state.connected === false && state.sessionId === null,
    )).toHaveLength(2);
  });

  it("coalesces hidden stop with pagehide disposal into one stop per run", async () => {
    const removeHandler = vi.fn(async () => {});
    const bridge = {
      sessionId: SESSION.sessionId,
      start: vi.fn(async () => SESSION),
      stop: vi.fn(async () => {}),
    };
    const onState = vi.fn();
    const controller = new OfficeBridgeLifecycleController(() => bridge, onState);
    const harness = eventTargetHarness();
    let visibilityHandler!: (message: { visibilityMode: unknown }) => void;
    controller.register(harness.target, {
      hiddenValue: "hidden",
      register: async (handler) => {
        visibilityHandler = handler;
        return removeHandler;
      },
    });

    await controller.start();
    await flush();
    visibilityHandler({ visibilityMode: "hidden" });
    harness.fire("pagehide");
    await flush();
    await flush();

    expect(bridge.stop).toHaveBeenCalledTimes(1);
    expect(onState.mock.calls.filter(([state]) =>
      state.connected === false && state.sessionId === null,
    )).toHaveLength(1);
    expect(removeHandler).toHaveBeenCalledTimes(1);
  });
});
