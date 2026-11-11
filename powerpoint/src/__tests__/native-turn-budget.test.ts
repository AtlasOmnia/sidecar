import { afterEach, beforeEach, expect, it, vi } from "vitest";
import { HermesAgent, TURN_TIMEOUT_MS, type AgentCallbacks } from "../agent";
vi.mock("../powerPointTools", () => ({ getPresentationContext: vi.fn(async () => "metadata"), executePowerPointToolParsed: vi.fn(), toolSchemas: [] }));
const native = { baseUrl: "https://localhost:47650/v1", model: "hermes-agent", maxIterations: 1 };
const options = { bridgeSessionId: "550e8400-e29b-41d4-a716-446655440000" };
const cb = (): AgentCallbacks => ({ onAssistantStart: vi.fn(), onAssistantDelta: vi.fn(), onAssistantFinal: vi.fn(), onActivity: vi.fn(), onStatus: vi.fn(), onModelCall: vi.fn() });
function delayedFetch(ms: number) {
  const mock = vi.fn((_url: string, init: RequestInit) => new Promise<Response>((resolve, reject) => {
    const timer = setTimeout(() => resolve(new Response(JSON.stringify({ choices: [{ message: { content: "Edited successfully" } }] }), { headers: { "content-type": "application/json" } })), ms);
    init.signal!.addEventListener("abort", () => { clearTimeout(timer); reject(init.signal!.reason); }, { once: true });
  }));
  vi.stubGlobal("fetch", mock); return mock;
}
beforeEach(() => vi.useFakeTimers());
afterEach(() => { vi.useRealTimers(); vi.unstubAllGlobals(); });
it("a native multi-tool turn longer than 20 seconds succeeds", async () => {
  const fetch = delayedFetch(25_000);
  const pending = new HermesAgent().run("edit", native, cb(), new AbortController().signal, options);
  const assertion = expect(pending).resolves.toMatchObject({ finalContent: "Edited successfully" });
  await vi.advanceTimersByTimeAsync(25_001); await assertion; expect(fetch).toHaveBeenCalledTimes(1);
});
it("native timeout occurs at its realistic turn limit once, never retries", async () => {
  const fetch = delayedFetch(1_000_000); const callbacks = cb();
  const pending = new HermesAgent().run("edit", native, callbacks, new AbortController().signal, options);
  let failure: unknown; pending.catch(error => { failure = error; });
  await vi.advanceTimersByTimeAsync(90_001);
  expect(failure, "native turn must survive the old 90-second budget").toBeUndefined();
  await vi.advanceTimersByTimeAsync(210_000);
  await expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
  expect(fetch).toHaveBeenCalledTimes(1); expect(callbacks.onModelCall).toHaveBeenCalledTimes(1);
});
it("Continue at the live 30-second checkpoint extends the in-flight native turn", async () => {
  delayedFetch(310_000); const callbacks = cb();
  callbacks.onCheckpoint = vi.fn().mockImplementation(async () => { await new Promise(resolve => setTimeout(resolve, 20_000)); return "continue"; });
  const pending = new HermesAgent().run("edit", native, callbacks, new AbortController().signal, options);
  const assertion = expect(pending).resolves.toMatchObject({ finalContent: "Edited successfully" });
  await vi.advanceTimersByTimeAsync(30_001); expect(callbacks.onCheckpoint).toHaveBeenCalledTimes(1);
  expect(callbacks.onStatus).toHaveBeenCalledWith(expect.stringContaining("Hermes is working"));
  await vi.advanceTimersByTimeAsync(280_000); await assertion;
});
it("Stop at the live checkpoint cancels the native call without retry", async () => {
  const fetch = delayedFetch(100_000); const callbacks = cb(); callbacks.onCheckpoint = vi.fn().mockResolvedValue("stop");
  const pending = new HermesAgent().run("edit", native, callbacks, new AbortController().signal, options);
  const assertion = expect(pending).resolves.toMatchObject({ stopReason: "checkpoint" });
  await vi.advanceTimersByTimeAsync(30_001); await assertion; expect(fetch).toHaveBeenCalledTimes(1);
});
it("an unanswered native checkpoint remains bounded and reports timeout, not Stop", async () => {
  const fetch = delayedFetch(1_000_000); const callbacks = cb();
  callbacks.onCheckpoint = vi.fn().mockImplementation(() => new Promise(() => {}));
  const pending = new HermesAgent().run("edit", native, callbacks, new AbortController().signal, options);
  const assertion = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
  await vi.advanceTimersByTimeAsync(300_001); await assertion;
  expect(callbacks.onCheckpoint).toHaveBeenCalledTimes(1); expect(fetch).toHaveBeenCalledTimes(1);
});
it("native network failures are not automatically replayed after possibly editing", async () => {
  const fetch = vi.fn().mockRejectedValue(new TypeError("network disconnected")); vi.stubGlobal("fetch", fetch);
  await expect(new HermesAgent().run("edit", native, cb(), new AbortController().signal, options)).rejects.toMatchObject({ toolError: { category: "network" } });
  expect(fetch).toHaveBeenCalledTimes(1);
});
it("non-native per-call limit stays 20 seconds and total turn limit stays 90 seconds", async () => {
  expect(TURN_TIMEOUT_MS).toBe(90_000); const fetch = delayedFetch(25_000);
  const pending = new HermesAgent().run("edit", { ...native, model: "direct" }, cb(), new AbortController().signal);
  const assertion = expect(pending).rejects.toMatchObject({ name: "TimeoutError" });
  await vi.advanceTimersByTimeAsync(20_001); await assertion; expect(fetch).toHaveBeenCalledTimes(1);
});
