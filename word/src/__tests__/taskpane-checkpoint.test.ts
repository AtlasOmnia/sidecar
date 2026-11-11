// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { waitForCheckpointDecision } from "../taskpane/checkpoint";

function checkpointElement(): HTMLElement {
  const element = document.createElement("div");
  element.innerHTML = '<button class="checkpoint-continue">Continue</button><button class="checkpoint-stop">Stop</button>';
  return element;
}

describe("task-pane checkpoint decision", () => {
  it("resolves Continue and removes its listeners", async () => {
    const element = checkpointElement();
    const signal = new AbortController().signal;
    const decision = waitForCheckpointDecision(element, signal);

    element.querySelector<HTMLButtonElement>(".checkpoint-continue")!.click();
    await expect(decision).resolves.toBe(true);

    const secondDecision = waitForCheckpointDecision(element, signal);
    element.querySelector<HTMLButtonElement>(".checkpoint-stop")!.click();
    await expect(secondDecision).resolves.toBe(false);
  });

  it("resolves Stop without aborting the turn", async () => {
    const controller = new AbortController();
    const element = checkpointElement();
    const decision = waitForCheckpointDecision(element, controller.signal);

    element.querySelector<HTMLButtonElement>(".checkpoint-stop")!.click();

    await expect(decision).resolves.toBe(false);
    expect(controller.signal.aborted).toBe(false);
  });

  it("rejects with the original user AbortError and cleans up", async () => {
    const controller = new AbortController();
    const reason = new DOMException("Stopped by the user.", "AbortError");
    const decision = waitForCheckpointDecision(checkpointElement(), controller.signal);

    controller.abort(reason);

    await expect(decision).rejects.toBe(reason);
  });

  it("rejects immediately when the deadline is already aborted", async () => {
    const controller = new AbortController();
    const reason = new DOMException("Turn timed out.", "TimeoutError");
    controller.abort(reason);

    await expect(waitForCheckpointDecision(checkpointElement(), controller.signal)).rejects.toBe(reason);
  });
});
