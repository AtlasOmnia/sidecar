// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { afterEach, beforeEach, expect, it, vi } from "vitest";
const harness = vi.hoisted(() => ({ grant: "synthetic-test-approval" as string | null, discover: vi.fn(), onState: undefined as ((state: unknown) => void) | undefined }));
vi.mock("../deskConnect", async (original) => ({ ...await original<typeof import("../deskConnect")>(), DeskConnector: class {
  get grant() { return harness.grant; }
  discover = harness.discover;
  clearGrant() { harness.grant = null; }
} }));
vi.mock("../officeBridgePowerPoint", () => ({ createPowerPointCompanionBridge: vi.fn() }));
vi.mock("../officeBridgeLifecycle", () => ({ OfficeBridgeLifecycleController: class {
  constructor(_factory: unknown, onState: (state: unknown) => void) { harness.onState = onState; }
  register() {}
  async stop() {}
  async start() { harness.onState!({ connected: true, sessionId: "test-session" }); }
} }));
beforeEach(async () => {
  vi.resetModules(); harness.grant = "synthetic-test-approval"; harness.discover.mockReset().mockResolvedValue("https://localhost:47650");
  document.documentElement.innerHTML = readFileSync("src/taskpane/taskpane.html", "utf8");
  let ready!: () => void;
  vi.stubGlobal("Office", { onReady: (callback: () => void) => { ready = callback; }, context: { platform: "Mac" } });
  await import("../taskpane/taskpane"); ready();
  await vi.waitFor(() => expect(document.getElementById("desk-connect")!.hidden).toBe(true));
});
afterEach(() => { vi.unstubAllGlobals(); });
it("a bridge failure with a grant shows connection loss, never approval or raw errors, and retries", async () => {
  harness.onState!({ connected: false, sessionId: null, lastError: "sensitive raw failure" });
  expect(document.getElementById("desk-connect")!.hidden).toBe(false);
  expect(document.getElementById("desk-connect-text")!.textContent).toBe("Lost the connection to Hermes on this computer.");
  expect(document.body.textContent).not.toContain("sensitive raw failure");
  const button = document.getElementById("desk-connect-button") as HTMLButtonElement;
  expect(button.textContent).toBe("Try again"); button.click();
  await vi.waitFor(() => expect(harness.discover).toHaveBeenCalledTimes(2));
});
it("a missing grant shows the approval message rather than connection loss", () => {
  harness.grant = null;
  harness.onState!({ connected: false, sessionId: null, lastError: "approval required" });
  expect(document.getElementById("desk-connect-text")!.textContent).toBe("This sidebar needs your OK in Hermes before it can chat.");
  expect(document.body.textContent).not.toContain("approval required");
});
