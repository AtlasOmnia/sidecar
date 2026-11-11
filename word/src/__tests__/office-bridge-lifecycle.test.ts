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
});
