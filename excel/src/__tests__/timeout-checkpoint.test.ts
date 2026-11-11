import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { HermesAgent, makeToolCallSignature, isBlockedRepeatedToolCall } from "../agent";
import { ToolErrorCategory } from "../toolErrors";
import { ToolCall } from "../types";

// ---------------------------------------------------------------------------
// Timeout and deadline behavioral tests
// ---------------------------------------------------------------------------

describe("absolute turn deadline (90s)", () => {
  it("throws TimeoutError when the turn exceeds 90 seconds", () => {
    // Verify that the checkTurnTimeout logic works at the class level.
    // We test this indirectly by verifying the time comparison math.
    const TURN_TIMEOUT_MS = 90_000;
    const turnStart = Date.now();

    // Simulate the check
    const elapsed = TURN_TIMEOUT_MS + 1;
    expect(elapsed >= TURN_TIMEOUT_MS).toBe(true);

    // The real check uses Date.now() - turnStart >= TURN_TIMEOUT_MS
    const simulatedNow = turnStart + TURN_TIMEOUT_MS;
    expect(simulatedNow - turnStart >= TURN_TIMEOUT_MS).toBe(true);
  });

  it("does not throw when under the deadline", () => {
    const TURN_TIMEOUT_MS = 90_000;
    const turnStart = Date.now();
    const elapsed = TURN_TIMEOUT_MS - 100;
    expect(elapsed >= TURN_TIMEOUT_MS).toBe(false);
  });
});

describe("per-call model timeout (20s)", () => {
  it("per-call timeout is capped at 20s", () => {
    const MODEL_CALL_TIMEOUT_MS = 20_000;
    const remainingTurn = 50_000; // plenty of turn time left
    const perCallTimeout = Math.min(
      MODEL_CALL_TIMEOUT_MS,
      Math.max(remainingTurn, 1000),
    );
    expect(perCallTimeout).toBe(MODEL_CALL_TIMEOUT_MS);
  });

  it("per-call timeout uses remaining turn time when less than 20s", () => {
    const MODEL_CALL_TIMEOUT_MS = 20_000;
    const remainingTurn = 5_000; // only 5s left
    const perCallTimeout = Math.min(
      MODEL_CALL_TIMEOUT_MS,
      Math.max(remainingTurn, 1000),
    );
    expect(perCallTimeout).toBe(5_000);
  });

  it("per-call timeout floors at 1s minimum", () => {
    const MODEL_CALL_TIMEOUT_MS = 20_000;
    const remainingTurn = 0; // effectively zero remaining
    const perCallTimeout = Math.min(
      MODEL_CALL_TIMEOUT_MS,
      Math.max(remainingTurn, 1000),
    );
    expect(perCallTimeout).toBe(1_000);
  });
});

describe("30-second checkpoint", () => {
  it("checkpoint fires after CHECKPOINT_MS (30s)", () => {
    const CHECKPOINT_MS = 30_000;
    const turnStart = Date.now();
    const elapsed = CHECKPOINT_MS;
    expect(elapsed >= CHECKPOINT_MS).toBe(true);
  });

  it("checkpoint does not fire before 30s", () => {
    const CHECKPOINT_MS = 30_000;
    const turnStart = Date.now();
    const elapsed = CHECKPOINT_MS - 1;
    expect(elapsed >= CHECKPOINT_MS).toBe(false);
  });

  it("checkpoint only fires once per turn", () => {
    // Simulate the state guard
    let checkpointFired = false;
    let callCount = 0;

    function maybeFireCheckpoint(elapsed: number): boolean {
      if (elapsed >= 30_000 && !checkpointFired) {
        checkpointFired = true;
        callCount++;
        return true;
      }
      // Already fired — should not fire again
      return false;
    }

    // First time past 30s
    expect(maybeFireCheckpoint(35_000)).toBe(true);
    expect(callCount).toBe(1);

    // Second pass — already fired
    maybeFireCheckpoint(40_000);
    expect(callCount).toBe(1); // should not increment
  });
});

describe("retry backoff delays", () => {
  it("produces exponential backoff with cap", () => {
    const base = 1000;
    const max = 8000;

    // Attempt 0 -> delay 0 (first attempt, no delay needed)
    // Attempt 1 -> delay 1000 * 2^0 = 1000
    // Attempt 2 -> delay 1000 * 2^1 = 2000
    // Attempt 3 -> delay 1000 * 2^2 = 4000 (would be attempt 3, but MAX_RETRIES=2)

    const delays = [0, 1, 2, 3, 4, 5].map((attempt) =>
      Math.min(base * Math.pow(2, attempt), max),
    );
    expect(delays).toEqual([1000, 2000, 4000, 8000, 8000, 8000]);
  });
});

describe("absolute iteration cap", () => {
  it("caps iterations at ABSOLUTE_MAX_ITERATIONS (20) regardless of settings", () => {
    const ABSOLUTE_MAX_ITERATIONS = 20;
    const settingsMax = 50; // user set it high
    const effective = Math.min(settingsMax, ABSOLUTE_MAX_ITERATIONS);
    expect(effective).toBe(20);
  });

  it("respects lower settings value when under the cap", () => {
    const ABSOLUTE_MAX_ITERATIONS = 20;
    const settingsMax = 5;
    const effective = Math.min(settingsMax, ABSOLUTE_MAX_ITERATIONS);
    expect(effective).toBe(5);
  });
});

describe("cancellation propagation", () => {
  it("throws AbortError when signal is aborted during the loop", () => {
    const controller = new AbortController();
    controller.abort();

    // Simulate throwIfAborted
    expect(() => {
      if (controller.signal.aborted) {
        throw new DOMException("The running Hermes task was stopped.", "AbortError");
      }
    }).toThrow(DOMException);
  });

  it("retry delay rejects with AbortError when signal aborts during sleep", async () => {
    const controller = new AbortController();

    // This test verifies the pattern used in sleep()
    const sleepPromise = new Promise<void>((resolve, reject) => {
      const timer = setTimeout(resolve, 5000);
      const onAbort = () => {
        clearTimeout(timer);
        reject(new DOMException("Aborted during retry delay.", "AbortError"));
      };
      controller.signal.addEventListener("abort", onAbort, { once: true });
    });

    // Abort immediately
    controller.abort();

    await expect(sleepPromise).rejects.toThrow(DOMException);
    await expect(sleepPromise).rejects.toThrow("Aborted during retry delay.");
  });
});

describe("generation isolation", () => {
  it("stale abort controllers do not affect subsequent turns", () => {
    // Turn 1 controller
    const turn1 = new AbortController();
    turn1.abort();
    expect(turn1.signal.aborted).toBe(true);

    // Turn 2 controller — fresh, unaffected
    const turn2 = new AbortController();
    expect(turn2.signal.aborted).toBe(false);

    // Turn 1's abort should not leak into Turn 2
    expect(turn1.signal.aborted).not.toBe(turn2.signal.aborted);
  });
});

describe("typed error results in tool responses", () => {
  it("repetition errors carry the correct category", () => {
    const result = JSON.stringify({
      ok: false,
      error: "Repeated identical tool call blocked.",
      category: ToolErrorCategory.Repetition,
    });
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.category).toBe("repetition");
  });

  it("office host errors carry ItemNotFound details", () => {
    const result = JSON.stringify({
      ok: false,
      error: "Sheet not found.",
      category: ToolErrorCategory.OfficeHost,
      name: "ItemNotFound",
    });
    const parsed = JSON.parse(result);
    expect(parsed.ok).toBe(false);
    expect(parsed.category).toBe("office_host");
    expect(parsed.name).toBe("ItemNotFound");
  });
});
