import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createLinkedDeadline, interruptibleSleep, raceWithSignal } from "../agentRuntime";

describe("agent runtime abort primitives", () => {
  beforeEach(() => vi.useFakeTimers());
  afterEach(() => vi.useRealTimers());

  it("preserves the parent AbortError reason", () => {
    const parent = new AbortController();
    const linked = createLinkedDeadline(parent.signal, 10_000, "deadline");
    const reason = new DOMException("Stopped by the user.", "AbortError");

    parent.abort(reason);

    expect(linked.signal.reason).toBe(reason);
    expect(linked.signal.reason.name).toBe("AbortError");
    linked.cleanup();
  });

  it("uses TimeoutError for a deadline and cleanup clears the timer", () => {
    const parent = new AbortController();
    const linked = createLinkedDeadline(parent.signal, 1_000, "Model call timed out.");

    vi.advanceTimersByTime(1_000);
    expect(linked.signal.reason).toMatchObject({ name: "TimeoutError", message: "Model call timed out." });
    linked.cleanup();
  });

  it("races a never-resolving operation against an abort signal", async () => {
    const controller = new AbortController();
    const pending = raceWithSignal(new Promise<string>(() => undefined), controller.signal);
    controller.abort(new DOMException("deadline", "TimeoutError"));

    await expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
  });

  it("observes a promise rejection even when the signal is already aborted", async () => {
    const controller = new AbortController();
    controller.abort(new DOMException("stopped", "AbortError"));
    let rejectPromise!: (reason: unknown) => void;
    const started = new Promise<never>((_resolve, reject) => {
      rejectPromise = reject;
    });
    const unhandled: unknown[] = [];
    const onUnhandled = (reason: unknown): void => { unhandled.push(reason); };
    process.on("unhandledRejection", onUnhandled);

    try {
      await expect(raceWithSignal(started, controller.signal)).rejects.toMatchObject({ name: "AbortError" });
      rejectPromise(new Error("late failure"));
      await Promise.resolve();
      await Promise.resolve();
      expect(unhandled).toEqual([]);
    } finally {
      process.off("unhandledRejection", onUnhandled);
    }
  });

  it("interruptibleSleep resolves and preserves an abort reason", async () => {
    const controller = new AbortController();
    const promise = interruptibleSleep(1_000, controller.signal);
    vi.advanceTimersByTime(1_000);
    await expect(promise).resolves.toBeUndefined();

    const stopped = new AbortController();
    const stoppedPromise = interruptibleSleep(1_000, stopped.signal);
    const reason = new DOMException("Stopped", "AbortError");
    stopped.abort(reason);
    await expect(stoppedPromise).rejects.toBe(reason);
  });
});
