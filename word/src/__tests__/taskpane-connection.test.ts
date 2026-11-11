// @vitest-environment jsdom
import { beforeEach, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";

const f = vi.hoisted(() => ({
  state: null as null | ((state: any) => void),
  ready: null as null | (() => void),
  grant: "test-approval", discover: vi.fn(async () => "https://localhost:12345"),
  start: vi.fn(async () => undefined), stop: vi.fn(async () => undefined),
}));
vi.mock("../deskConnect", () => ({
  DESK_COMMAND: "/sidecar-word", ApprovalRequiredError: class extends Error {},
  DeskConnector: class {
    get grant() { return f.grant; }
    discover = f.discover;
    clearGrant() { f.grant = ""; }
  },
}));
vi.mock("../officeBridgeLifecycle", () => ({
  OfficeBridgeLifecycleController: class {
    constructor(_factory: unknown, state: any) { f.state = state; }
    start = f.start; stop = f.stop; register() {}
  },
}));

beforeEach(async () => {
  vi.resetModules();
  f.grant = "test-approval";
  f.discover.mockClear(); f.start.mockReset().mockResolvedValue(undefined);
  document.documentElement.innerHTML = readFileSync("src/taskpane/taskpane.html", "utf8");
  vi.stubGlobal("Office", { onReady: (fn: () => void) => { f.ready = fn; },
    context: { requirements: { isSetSupported: () => true } } });
  await import("../taskpane/taskpane");
});

it("shows connection loss, never raw error detail or an approval prompt, and retries", async () => {
  f.ready!();
  await vi.waitFor(() => expect(f.start).toHaveBeenCalled());
  f.state!({ connected: false, sessionId: null, lastError: "https://localhost/?grant=do-not-display" });
  expect(document.getElementById("desk-connect")!.hidden).toBe(false);
  expect(document.getElementById("desk-connect-text")!.textContent).toBe("Lost the connection to Hermes on this computer.");
  expect(document.body.textContent).not.toContain("do-not-display");
  const button = document.getElementById("desk-connect-button") as HTMLButtonElement;
  expect(button.textContent).toBe("Try again");
  button.click();
  await vi.waitFor(() => expect(f.discover).toHaveBeenCalledTimes(2));
});

it("asks for approval only when no grant exists", () => {
  f.grant = "";
  f.state!({ connected: false, sessionId: null, lastError: "connection failed" });
  expect(document.getElementById("desk-connect-text")!.textContent).toBe("This sidebar needs your OK in Hermes before it can chat.");
});

it("keeps generic connect failures secret-free and distinct from approval failures", async () => {
  f.start.mockRejectedValueOnce(new Error("https://localhost/?grant=do-not-display"));
  f.ready!();
  await vi.waitFor(() => expect(document.getElementById("desk-connect-text")!.textContent)
    .toBe("Lost the connection to Hermes on this computer."));
  expect(document.body.textContent).not.toContain("do-not-display");
});

it("clears rejected approval and shows the approval message", async () => {
  const { ApprovalRequiredError } = await import("../deskConnect");
  f.start.mockRejectedValueOnce(new ApprovalRequiredError());
  f.ready!();
  await vi.waitFor(() => expect(document.getElementById("desk-connect-text")!.textContent)
    .toBe("This sidebar needs your OK in Hermes before it can chat."));
  expect(f.grant).toBe("");
});

it("provides every literal byId target in the HTML", () => {
  const source = readFileSync("src/taskpane/taskpane.ts", "utf8");
  const ids = [...source.matchAll(/byId(?:<[^>]+>)?\("([^"]+)"\)/g)].map(match => match[1]);
  expect(ids.length).toBeGreaterThan(0);
  expect(ids.filter(id => !document.getElementById(id))).toEqual([]);
});
