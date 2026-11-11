import { afterEach, expect, it, vi } from "vitest";
import { HermesAgent, type AgentCallbacks } from "../agent";
vi.mock("../powerPointTools", () => ({ getPresentationContext: vi.fn(async () => "metadata"), executePowerPointToolParsed: vi.fn(), toolSchemas: [] }));
const native = { baseUrl: "https://localhost:47650/v1", model: "hermes-agent", maxIterations: 1 };
const options = { bridgeSessionId: "550e8400-e29b-41d4-a716-446655440000" };
const callbacks = (): AgentCallbacks => ({ onAssistantStart: vi.fn(), onAssistantDelta: vi.fn(), onAssistantFinal: vi.fn(), onActivity: vi.fn(), onStatus: vi.fn() });
afterEach(() => { vi.unstubAllGlobals(); vi.useRealTimers(); });
it.each(["empty", "done-only", "no-body"])("never replays a dispatched native request after %s response", async (kind) => {
  const response = kind === "no-body" ? new Response(null) : new Response(kind === "empty" ? "" : "data: [DONE]\n\n", { headers: { "content-type": "text/event-stream" } });
  const fetch = vi.fn().mockResolvedValueOnce(response).mockResolvedValue(new Response(JSON.stringify({ choices: [{ message: { content: "Already done" } }] }), { headers: { "content-type": "application/json" } }));
  vi.stubGlobal("fetch", fetch);
  await expect(new HermesAgent().run("edit", native, callbacks(), new AbortController().signal, options)).rejects.toMatchObject({ toolError: { category: "protocol" } });
  expect(fetch).toHaveBeenCalledTimes(1);
});
it("keeps one native request un-aborted through tool execution and a six-second response delay", async () => {
  vi.useFakeTimers();
  let toolFinished = false;
  const fetch = vi.fn((_url: string, init: RequestInit) => new Promise<Response>((resolve, reject) => {
    setTimeout(() => { toolFinished = true; }, 4_000);
    setTimeout(() => resolve(new Response('data: {"choices":[{"delta":{"content":"Edited successfully"}}]}\n\ndata: [DONE]\n\n')), 6_000);
    init.signal!.addEventListener("abort", () => reject(init.signal!.reason), { once: true });
  }));
  vi.stubGlobal("fetch", fetch);
  const pending = new HermesAgent().run("edit", native, callbacks(), new AbortController().signal, options);
  const assertion = expect(pending).resolves.toMatchObject({ finalContent: "Edited successfully" });
  await vi.advanceTimersByTimeAsync(5_500);
  expect(toolFinished).toBe(true);
  expect(fetch).toHaveBeenCalledTimes(1);
  expect(fetch.mock.calls[0][1].signal!.aborted).toBe(false);
  await vi.advanceTimersByTimeAsync(501);
  await assertion;
  expect(fetch).toHaveBeenCalledTimes(1);
  expect(fetch.mock.calls[0][1].signal!.aborted).toBe(false);
});
