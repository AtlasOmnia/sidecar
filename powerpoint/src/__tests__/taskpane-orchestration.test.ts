import { describe, expect, it } from "vitest";
import { CheckpointController } from "../checkpoint";

describe("task-pane checkpoint coordination", () => {
  it("resolves exactly once when the user chooses Continue", async () => {
    const controller = new CheckpointController();
    const signal = new AbortController();
    const pending = controller.begin(signal.signal);
    expect(controller.isPending).toBe(true);
    controller.choose("continue");
    expect(await pending).toBe("continue");
    controller.choose("stop");
    expect(controller.isPending).toBe(false);
  });

  it("settles a pending prompt as Stop when its deadline aborts", async () => {
    const controller = new CheckpointController();
    const abort = new AbortController();
    const pending = controller.begin(abort.signal);
    abort.abort(new DOMException("deadline", "TimeoutError"));
    await expect(pending).resolves.toBe("stop");
    expect(controller.isPending).toBe(false);
  });

  it("does not leave a stale pending resolver after cancellation", async () => {
    const controller = new CheckpointController();
    const abort = new AbortController();
    const pending = controller.begin(abort.signal);
    controller.cancel();
    await expect(pending).resolves.toBe("stop");
    expect(controller.begin(new AbortController().signal)).not.toBe(pending);
    controller.choose("continue");
  });
});
