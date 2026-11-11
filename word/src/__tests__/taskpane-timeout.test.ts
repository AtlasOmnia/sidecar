// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

describe("task-pane timeout reporting", () => {
  it("keeps sanitized selected-budget detail and the possible-action warning while Stop stays distinct", async () => {
    vi.stubGlobal("Office", {
      onReady: vi.fn(),
      addin: undefined,
      context: { requirements: { isSetSupported: () => true } },
    });
    const taskpane = await import("../taskpane/taskpane");
    const module = taskpane as unknown as {
      formatTaskpaneTimeoutMessage?: (error: unknown) => string;
      formatTaskpaneStoppedMessage?: (assistantBuffer: string) => string;
    };

    const timeoutMessage = module.formatTaskpaneTimeoutMessage?.(
      new DOMException(
        "Turn exceeded the 210-second absolute time limit. Document actions may have occurred; review the document before retrying. api_key=secret",
        "TimeoutError",
      ),
    );
    const stoppedMessage = module.formatTaskpaneStoppedMessage?.("partial result");

    expect(typeof timeoutMessage).toBe("string");
    expect(typeof stoppedMessage).toBe("string");
    expect(timeoutMessage as string).toContain("210-second absolute time limit");
    expect(timeoutMessage as string).toContain("Document actions may have occurred; review the document before retrying.");
    expect(timeoutMessage as string).not.toContain("api_key=secret");
    expect(stoppedMessage as string).toContain("Stopped.");
    expect(stoppedMessage as string).not.toContain("210-second absolute time limit");
  });
});
