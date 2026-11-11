import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { HermesAgent } from "../agent";
import { getPresentationContext } from "../powerPointTools";

const execute = vi.fn();
vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn().mockResolvedValue("deck metadata"),
  executePowerPointToolParsed: (...args: unknown[]) => execute(...args),
  toolSchemas: [{ type: "function", function: { name: "test_tool", description: "test", parameters: { type: "object", properties: { value: { type: "string" }, options: { type: "object", properties: { a: { type: "number" }, b: { type: "number" } }, required: ["a", "b"], additionalProperties: false } }, required: ["value"], additionalProperties: false } } }],
}));
const settings = { baseUrl: "http://test.local/v1", model: "test", maxIterations: 20 };
const cb = () => ({ onAssistantStart: vi.fn(), onAssistantDelta: vi.fn(), onAssistantFinal: vi.fn(), onActivity: vi.fn(), onStatus: vi.fn(), onModelCall: vi.fn(), onToolResult: vi.fn() });
function jsonResponse(content: string, toolCalls?: unknown[]): Response { return new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content, ...(toolCalls ? { tool_calls: toolCalls } : {}) } }] }), { status: 200, headers: { "content-type": "application/json" } }); }
function call(args: string, id = "call-1") { return { id, type: "function", function: { name: "test_tool", arguments: args } }; }

describe("PowerPoint orchestration end to end", () => {
  beforeEach(() => {
    vi.mocked(getPresentationContext).mockResolvedValue("deck metadata");
    execute.mockReset().mockResolvedValue(JSON.stringify({ ok: true }));
  });
  afterEach(() => { vi.restoreAllMocks(); vi.unstubAllGlobals(); });
  it("streams normal assistant text", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue(new Response("data: {\"choices\":[{\"delta\":{\"content\":\"hello\"}}]}\n\ndata: [DONE]\n\n", { status: 200, headers: { "content-type": "text/event-stream" } })));
    const callbacks = cb(); const result = await new HermesAgent().run("hello", settings, callbacks, new AbortController().signal);
    expect(result.finalContent).toBe("hello"); expect(callbacks.onAssistantDelta).toHaveBeenCalledWith("hello");
    expect(callbacks.onModelCall).toHaveBeenCalledWith(expect.objectContaining({ attempt: 1, statusCode: 200, outcome: "success" }));
  });
  it("executes only validated parsed arguments and then continues", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(jsonResponse("", [call('{"value":"ok"}')])).mockResolvedValueOnce(jsonResponse("finished")));
    const result = await new HermesAgent().run("format", settings, cb(), new AbortController().signal);
    expect(execute).toHaveBeenCalledWith("test_tool", { value: "ok" }); expect(result.finalContent).toBe("finished");
  });
  it("rejects nested junk before Office execution", async () => {
    vi.stubGlobal("fetch", vi.fn().mockResolvedValueOnce(jsonResponse("", [call('{"slideIndex":1,"font":{"secret":"sentinel"}}')])).mockResolvedValueOnce(jsonResponse("corrected")));
    const result = await new HermesAgent().run("format", settings, cb(), new AbortController().signal);
    expect(execute).not.toHaveBeenCalled(); expect(result.finalContent).toBe("corrected");
  });
  it("stops a pending fetch as AbortError rather than a network error", async () => {
    const controller = new AbortController(); vi.stubGlobal("fetch", vi.fn().mockImplementation(() => new Promise((_resolve, reject) => controller.signal.addEventListener("abort", () => reject(controller.signal.reason)))));
    const promise = new HermesAgent().run("stop", settings, cb(), controller.signal); controller.abort(new DOMException("stop", "AbortError"));
    await expect(promise).rejects.toMatchObject({ name: "AbortError" });
  });
  it("blocks the real third equivalent reordered and whitespace tool call", async () => {
    const same = [
      jsonResponse("", [call('{"value":"same","options":{"b":2,"a":1}}')]),
      jsonResponse("", [call('{ "options" : { "a" : 1, "b" : 2 }, "value" : "same" }', "two")]),
      jsonResponse("", [call('{"options":{"a":1,"b":2},"value":"same"}', "three")]),
      jsonResponse("finished")
    ];
    vi.stubGlobal("fetch", vi.fn().mockImplementation(() => Promise.resolve(same.shift()!)));
    const callbacks = cb();
    const result = await new HermesAgent().run("repeat", settings, callbacks, new AbortController().signal);
    expect(execute).toHaveBeenCalledTimes(2);
    expect(result.finalContent).toBe("finished");
    expect(JSON.stringify(callbacks.onToolResult.mock.calls)).toContain("repetition");
  });

  it("enforces the actual twenty iteration cap when settings request fifty", async () => {
    const calls = Array.from({ length: 21 }, (_, index) => jsonResponse("", [call(`{"value":"${index}"}`, `call-${index}`)]));
    vi.stubGlobal("fetch", vi.fn().mockImplementation(() => Promise.resolve(calls.shift()!)));
    const callbacks = cb();
    const result = await new HermesAgent().run("loop", { ...settings, maxIterations: 50 }, callbacks, new AbortController().signal);
    expect(global.fetch).toHaveBeenCalledTimes(20);
    expect(result.stopReason).toBe("iteration_limit");
    expect(result.finalContent).toContain("maximum of 20");
  });


});
