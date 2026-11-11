import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { HermesAgent, type AgentCallbacks } from "../agent";
import { executePowerPointToolParsed, getPresentationContext } from "../powerPointTools";

vi.mock("../powerPointTools", () => ({
  getPresentationContext: vi.fn(),
  executePowerPointToolParsed: vi.fn(),
  toolSchemas: [{
    type: "function",
    function: {
      name: "inspect_slides",
      description: "Inspect slides",
      parameters: { type: "object", properties: {}, additionalProperties: false },
    },
  }],
}));

const nativeSettings = { baseUrl: "https://localhost:47650/v1", model: "hermes-agent", maxIterations: 1, requestHeaders: { "X-Hermes-Session-Key": "synthetic-session-key", "X-Hermes-Pane-Id": "synthetic-pane" } };
const directSettings = { baseUrl: "https://remote.example/v1", model: "hermes-agent", maxIterations: 1 };
const callbacks = (): AgentCallbacks => ({
  onAssistantStart: vi.fn(),
  onAssistantDelta: vi.fn(),
  onAssistantFinal: vi.fn(),
  onActivity: vi.fn(),
  onStatus: vi.fn(),
});
const finalResponse = (content: string): Response => new Response(JSON.stringify({ choices: [{ message: { role: "assistant", content } }] }), { status: 200, headers: { "content-type": "application/json" } });
const toolResponse = (content: string | null, xml = false): Response => new Response(JSON.stringify({ choices: [{ message: {
  role: "assistant",
  content: xml ? `<tool_call>{"name":"inspect_slides","arguments":{}}</tool_call>${content ?? ""}` : content,
  tool_calls: xml ? undefined : [{ id: "call-1", type: "function", function: { name: "inspect_slides", arguments: "{}" } }],
} }] }), { status: 200, headers: { "content-type": "application/json" } });

beforeEach(() => {
  vi.mocked(getPresentationContext).mockResolvedValue("presentation metadata");
  vi.mocked(executePowerPointToolParsed).mockResolvedValue(JSON.stringify({ ok: true }));
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("native Hermes request context", () => {
  it("sends the exact bridge session header and omits identity from the body and messages", async () => {
    const fetchMock = vi.fn().mockResolvedValue(finalResponse("done"));
    vi.stubGlobal("fetch", fetchMock);
    const sessionId = "550e8400-e29b-41d4-a716-446655440000";

    await new HermesAgent().run("inspect", nativeSettings, callbacks(), new AbortController().signal, { bridgeSessionId: sessionId });

    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(url).toBe("https://localhost:47650/v1/chat/completions");
    expect((init.headers as Record<string, string>)["X-Hermes-Session-Id"]).toBe(sessionId);
    expect(init.headers).toMatchObject(nativeSettings.requestHeaders);
    expect(init.headers).not.toHaveProperty("Authorization");
    const body = JSON.parse(String(init.body));
    expect(body).toEqual(expect.objectContaining({ model: "hermes-agent", stream: true, temperature: 0.2 }));
    expect(body).not.toHaveProperty("tools");
    expect(body).not.toHaveProperty("tool_choice");
    expect(JSON.stringify(body)).not.toContain(sessionId);
  });

  it("fails closed without a stream-false fallback or local tools after an empty native response", async () => {
    const fetchMock = vi.fn()
      .mockResolvedValueOnce(new Response(null, { status: 200, headers: { "content-type": "text/event-stream" } }))
      .mockResolvedValueOnce(finalResponse("fallback"));
    vi.stubGlobal("fetch", fetchMock);
    const agent = new HermesAgent();
    const cb = callbacks();

    await expect(agent.run("inspect", nativeSettings, cb, new AbortController().signal, { bridgeSessionId: "550e8400-e29b-41d4-a716-446655440000" }))
      .rejects.toMatchObject({ toolError: { category: "protocol" } });

    expect(fetchMock).toHaveBeenCalledTimes(1);
    const body = JSON.parse(String((fetchMock.mock.calls[0][1] as RequestInit).body));
    expect(body.stream).toBe(true);
    expect(body).not.toHaveProperty("tools");
    expect(body).not.toHaveProperty("tool_choice");
    expect(executePowerPointToolParsed).not.toHaveBeenCalled();
    expect(cb.onAssistantFinal).not.toHaveBeenCalled();
    expect(agent.history).toEqual([]);
  });

  it("fails closed before fetch when the native bridge session is absent or invalid", async () => {
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    const agent = new HermesAgent();

    await expect(agent.run("inspect", nativeSettings, callbacks(), new AbortController().signal)).rejects.toMatchObject({ toolError: { category: "validation" } });
    await expect(agent.run("inspect", nativeSettings, callbacks(), new AbortController().signal, { bridgeSessionId: "not-a-uuid" })).rejects.toMatchObject({ toolError: { category: "validation" } });
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("never executes server tool_calls or XML tool-call output locally", async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(toolResponse(null)).mockResolvedValueOnce(toolResponse("server final", true));
    vi.stubGlobal("fetch", fetchMock);
    const agent = new HermesAgent();
    const cb = callbacks();

    const first = await agent.run("inspect", nativeSettings, cb, new AbortController().signal, { bridgeSessionId: "550e8400-e29b-41d4-a716-446655440000" });
    expect(first.finalContent).toBe("");
    const second = await agent.run("inspect", nativeSettings, cb, new AbortController().signal, { bridgeSessionId: "550e8400-e29b-41d4-a716-446655440000" });
    expect(second.finalContent).toBe("server final");
    expect(executePowerPointToolParsed).not.toHaveBeenCalled();
  });

  it("preserves direct-model schemas and local tool execution", async () => {
    const fetchMock = vi.fn().mockResolvedValueOnce(toolResponse(null)).mockResolvedValueOnce(finalResponse("local result"));
    vi.stubGlobal("fetch", fetchMock);

    await new HermesAgent().run("inspect", directSettings, callbacks(), new AbortController().signal);

    const body = JSON.parse(String((fetchMock.mock.calls[0][1] as RequestInit).body));
    expect(body.tools).toBeDefined();
    expect(body.tool_choice).toBe("auto");
    expect(executePowerPointToolParsed).toHaveBeenCalledWith("inspect_slides", {});
  });
});
