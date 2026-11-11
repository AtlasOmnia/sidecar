import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, type AgentCallbacks } from "../agent";
import { getPresentationContext, executePowerPointToolParsed } from "../powerPointTools";

vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn(),
  executePowerPointToolParsed: vi.fn(),
  toolSchemas: [{ type: "function", function: { name: "test_tool", description: "test", parameters: { type: "object", properties: { value: { type: "string" } }, required: ["value"], additionalProperties: false } } }]
}));

const settings = { baseUrl: "http://test.local/v1", model: "test", maxIterations: 3 };
const callbacks = (): AgentCallbacks => ({ onAssistantStart: vi.fn(), onAssistantDelta: vi.fn(), onAssistantFinal: vi.fn(), onActivity: vi.fn(), onStatus: vi.fn() });
const finalResponse = () => new Response("data: {\"choices\":[{\"delta\":{\"content\":\"done\"}}]}\n\ndata: [DONE]\n\n", { status: 200, headers: { "content-type": "text/event-stream" } });
const toolResponse = () => new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content: "", tool_calls: [{ id: "tool-1", type: "function", function: { name: "test_tool", arguments: '{"value":"x"}' } }] } }] }), { status: 200, headers: { "content-type": "application/json" } });

beforeEach(() => {
  vi.mocked(getPresentationContext).mockResolvedValue("metadata");
  vi.mocked(executePowerPointToolParsed).mockResolvedValue(JSON.stringify({ ok: true }));
});
afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); vi.useRealTimers(); });

describe("whole-turn containment and generation isolation", () => {
  it("bounds a non-cooperative presentation-context promise", async () => {
    vi.useFakeTimers();
    vi.mocked(getPresentationContext).mockReturnValue(new Promise(() => {}));
    const promise = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    const assertion = expect(promise).rejects.toMatchObject({ name: "TimeoutError", toolError: { category: "timeout" } });
    await vi.advanceTimersByTimeAsync(90_001);
    await assertion;
  });

  it("bounds a non-cooperative model fetch at the model deadline", async () => {
    vi.useFakeTimers();
    vi.stubGlobal("fetch", vi.fn(() => new Promise(() => {})));
    const promise = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    const assertion = expect(promise).rejects.toMatchObject({ name: "TimeoutError", toolError: { category: "timeout" } });
    await vi.advanceTimersByTimeAsync(20_001);
    await assertion;
  });

  it("bounds a non-cooperative checkpoint callback at the turn deadline", async () => {
    vi.useFakeTimers({ now: 0 });
    vi.mocked(getPresentationContext).mockImplementation(async () => { vi.setSystemTime(30_000); return "metadata"; });
    const cb = callbacks();
    cb.onCheckpoint = vi.fn().mockImplementation(() => new Promise<"continue" | "stop">(() => {}));
    const promise = new HermesAgent().run("hello", settings, cb, new AbortController().signal);
    const assertion = expect(promise).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(90_001);
    await assertion;
  });

  it("bounds a non-cooperative Office mutation without retrying or committing", async () => {
    vi.useFakeTimers();
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(toolResponse()));
    vi.mocked(executePowerPointToolParsed).mockReturnValue(new Promise(() => {}));
    const agent = new HermesAgent();
    const promise = agent.run("edit", settings, callbacks(), new AbortController().signal);
    const assertion = expect(promise).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(90_001);
    await assertion;
    expect(agent.history).toEqual([]);
    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  it("preserves arbitrary checkpoint errors instead of classifying TypeError as network", async () => {
    vi.useFakeTimers({ now: 0 });
    vi.mocked(getPresentationContext).mockImplementation(async () => { vi.setSystemTime(30_000); return "metadata"; });
    const cb = callbacks();
    const failure = new TypeError("checkpoint programming bug");
    cb.onCheckpoint = vi.fn().mockRejectedValue(failure);
    await expect(new HermesAgent().run("hello", settings, cb, new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" }, message: "An internal Hermes error occurred." });
  });

  it("preserves arbitrary tool errors instead of classifying TypeError as network", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(toolResponse()));
    vi.mocked(executePowerPointToolParsed).mockRejectedValue(new TypeError("tool programming bug"));
    await expect(new HermesAgent().run("edit", settings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "internal" }, message: "An internal Hermes error occurred." });
  });

  it("supports checkpoint Continue and Stop with distinct outcomes", async () => {
    vi.useFakeTimers({ now: 0 });
    vi.mocked(getPresentationContext).mockImplementation(async () => { vi.setSystemTime(30_000); return "metadata"; });
    const stop = callbacks(); stop.onCheckpoint = vi.fn().mockResolvedValue("stop");
    const stopped = await new HermesAgent().run("hello", settings, stop, new AbortController().signal);
    expect(stopped.stopReason).toBe("checkpoint");
    expect(stopped.finalContent).toContain("checkpoint");
    vi.setSystemTime(0);
    const cont = callbacks(); cont.onCheckpoint = vi.fn().mockResolvedValue("continue");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    await expect(new HermesAgent().run("hello", settings, cont, new AbortController().signal)).resolves.toMatchObject({ finalContent: "done" });
    expect(cont.onCheckpoint).toHaveBeenCalledTimes(1);
  });

  it("rejects concurrent runs and prevents a reset stale generation from committing", async () => {
    let resolveContext!: (value: string) => void;
    vi.mocked(getPresentationContext).mockReturnValue(new Promise((resolve) => { resolveContext = resolve; }));
    const agent = new HermesAgent();
    const firstCallbacks = callbacks();
    const first = agent.run("first", settings, firstCallbacks, new AbortController().signal);
    await expect(agent.run("second", settings, callbacks(), new AbortController().signal)).rejects.toThrow(/already running/);
    agent.reset();
    resolveContext("metadata");
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    await expect(first).rejects.toMatchObject({ name: "AbortError" });
    expect(agent.history).toEqual([]);
    expect(firstCallbacks.onAssistantFinal).not.toHaveBeenCalled();
  });

  it("does not start a retry when the remaining turn budget is too small", async () => {
    vi.useFakeTimers({ now: 0 });
    vi.mocked(getPresentationContext).mockImplementation(async () => { vi.setSystemTime(89_900); return "metadata"; });
    const fetchMock = vi.fn().mockRejectedValue(new TypeError("temporary network"));
    vi.stubGlobal("fetch", fetchMock);
    const promise = new HermesAgent().run("hello", settings, callbacks(), new AbortController().signal);
    const assertion = expect(promise).rejects.toMatchObject({ name: "TimeoutError", toolError: { category: "timeout" } });
    await vi.runAllTimersAsync();
    await assertion;
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });


  it("does not commit history when the final callback fails", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(finalResponse()));
    const cb = callbacks();
    cb.onAssistantFinal = vi.fn(() => { throw new Error("render failed"); });
    const agent = new HermesAgent();
    await expect(agent.run("hello", settings, cb, new AbortController().signal)).rejects.toThrow("An internal Hermes error occurred.");
    expect(agent.history).toEqual([]);
  });
});
