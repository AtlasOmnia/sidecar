import { describe, expect, it, vi } from "vitest";
import { HermesAgent } from "../agent";
import {
  MAX_PROTOCOL_ARGUMENT_BYTES,
  MAX_PROTOCOL_CONTENT_BYTES,
  MAX_PROTOCOL_TOOL_CALL_BYTES,
  MAX_PROTOCOL_TOOL_CALLS,
  isWorkbookPlanProtocolCandidate as isWorkbookPlanProtocolCandidateRaw,
  normalizeWorkbookPlanResponse as normalizeWorkbookPlanResponseRaw,
} from "../workbookPlanProtocol";

const validPlan = () => ({
  version: 1,
  outcome: "execute" as const,
  summary: "Write one value",
  operations: [
    {
      kind: "write_values" as const,
      sheet: "Sheet1",
      address: "A1",
      values: [["ok"]],
    },
  ],
  verification: [
    {
      kind: "range_values_equal" as const,
      sheet: "Sheet1",
      address: "A1",
      expectedValues: [["ok"]],
    },
  ],
});

const openAiCall = (argumentsValue: unknown, overrides: Record<string, unknown> = {}) => ({
  id: "call_plan_1",
  type: "function",
  function: {
    name: "excel_execute_plan",
    arguments: JSON.stringify(argumentsValue),
  },
  ...overrides,
});

type ProtocolInput = {
  readonly content: string | null;
  readonly toolCalls: string;
};

const normalizeWorkbookPlanResponse = (input: ProtocolInput) =>
  normalizeWorkbookPlanResponseRaw(input.content, input.toolCalls);

const isWorkbookPlanProtocolCandidate = (input: ProtocolInput) =>
  isWorkbookPlanProtocolCandidateRaw(input.content, input.toolCalls);

const openAi = (toolCalls: unknown[], content: string | null = null): ProtocolInput => ({
  content,
  toolCalls: JSON.stringify(toolCalls),
});

const xml = (payload: unknown, content = "") =>
  openAi([], `${content}<tool_call>${JSON.stringify({ name: "excel_execute_plan", arguments: payload })}</tool_call>`);

const primitiveXml = (name: string, argumentsValue: unknown) =>
  openAi([], `<tool_call>${JSON.stringify({ name, arguments: argumentsValue })}</tool_call>`);

const expectRejected = async (input: ProtocolInput) => {
  const result = await normalizeWorkbookPlanResponse(input);
  expect(result.ok).toBe(false);
  if (result.ok !== false) throw new Error("Expected protocol rejection");
  expect(result.error).toEqual({
    code: "protocol_rejected",
    message: "Invalid workbook plan protocol envelope",
  });
  expect(JSON.stringify(result)).not.toMatch(/Sheet1|A1|secret|Write one|updated/i);
};

describe("WorkbookPlan protocol normalization", () => {
  it("requires serialized tool-call transport and rejects former live-object positions without traps", async () => {
    const trapCounts = { get: 0, ownKeys: 0, descriptor: 0, prototype: 0 };
    const hostile = new Proxy([openAiCall(validPlan())], {
      get() {
        trapCounts.get += 1;
        throw new Error("get trap");
      },
      ownKeys() {
        trapCounts.ownKeys += 1;
        throw new Error("ownKeys trap");
      },
      getOwnPropertyDescriptor() {
        trapCounts.descriptor += 1;
        throw new Error("descriptor trap");
      },
      getPrototypeOf() {
        trapCounts.prototype += 1;
        throw new Error("prototype trap");
      },
    });

    await expectRejected({ content: "", toolCalls: hostile as unknown as string });
    expect(trapCounts).toEqual({ get: 0, ownKeys: 0, descriptor: 0, prototype: 0 });

    const serialized = JSON.stringify([openAiCall(validPlan())]);
    const accepted = await normalizeWorkbookPlanResponse({ content: "", toolCalls: serialized });
    expect(accepted.ok).toBe(true);
    if (!accepted.ok || accepted.kind !== "plan") throw new Error("Expected serialized plan");
    expect(accepted.plan).toEqual(validPlan());
  });

  it("rejects former outer input proxies in either primitive argument without traps", async () => {
    const trapCounts = { descriptor: 0, prototype: 0 };
    const formerInput = new Proxy(
      { content: "", toolCalls: JSON.stringify([]) },
      {
        getOwnPropertyDescriptor(target, key) {
          trapCounts.descriptor += 1;
          return Reflect.getOwnPropertyDescriptor(target, key);
        },
        getPrototypeOf(target) {
          trapCounts.prototype += 1;
          return Reflect.getPrototypeOf(target);
        },
      },
    );

    await expect(normalizeWorkbookPlanResponseRaw(formerInput, JSON.stringify([]))).resolves.toMatchObject({
      ok: false,
    });
    expect(isWorkbookPlanProtocolCandidateRaw(formerInput, JSON.stringify([]))).toBe(true);
    await expect(normalizeWorkbookPlanResponseRaw("", formerInput)).resolves.toMatchObject({
      ok: false,
    });
    expect(isWorkbookPlanProtocolCandidateRaw("", formerInput)).toBe(true);
    expect(trapCounts).toEqual({ descriptor: 0, prototype: 0 });
  });

  it("accepts exactly one OpenAI plan tool call with a direct payload", async () => {
    const result = await normalizeWorkbookPlanResponse(openAi([openAiCall(validPlan())]));
    expect(result.ok).toBe(true);
    if (!result.ok || result.kind !== "plan") throw new Error("Expected normalized plan");
    expect(result.transport).toBe("openai-tool-call");
    expect(result.plan).toEqual(validPlan());
    expect(Object.isFrozen(result)).toBe(true);
    expect(Object.isFrozen(result.plan)).toBe(true);
    if (result.plan.outcome !== "execute") throw new Error("Expected execute plan");
    expect(Object.isFrozen(result.plan.operations)).toBe(true);
    expect(Object.isFrozen(result.plan.operations[0])).toBe(true);
  });

  it("accepts one bounded Hermes XML fallback with the same direct object payload", async () => {
    const input = xml(validPlan());
    expect(isWorkbookPlanProtocolCandidate(input)).toBe(true);
    const result = await normalizeWorkbookPlanResponse(input);
    expect(result.ok).toBe(true);
    if (!result.ok || result.kind !== "plan") throw new Error("Expected normalized plan");
    expect(result.transport).toBe("xml-fallback");
    expect(result.plan).toEqual(validPlan());
  });


  it("returns a fresh immutable result and plan for each normalization", async () => {
    const input = openAi([openAiCall(validPlan())]);
    const first = await normalizeWorkbookPlanResponse(input);
    const second = await normalizeWorkbookPlanResponse(input);
    expect(first).not.toBe(second);
    if (!first.ok || !second.ok || first.kind !== "plan" || second.kind !== "plan") {
      throw new Error("Expected normalized plans");
    }
    if (first.plan.outcome !== "execute" || second.plan.outcome !== "execute") {
      throw new Error("Expected execute plans");
    }
    expect(first.plan).not.toBe(second.plan);
    expect(first.plan.operations).not.toBe(second.plan.operations);
    expect(() => {
      (first.plan as Record<string, unknown>).outcome = "refuse";
    }).toThrow();
  });

  it("preserves primitive compatibility when no plan envelope is present", async () => {
    const result = await normalizeWorkbookPlanResponse(
      openAi([
        {
          id: "call_primitive",
          type: "function",
          function: { name: "read_range", arguments: '{"sheet":"Sheet1","address":"A1"}' },
        },
      ], "I need to inspect the workbook."),
    );
    expect(result).toEqual({ ok: true, kind: "none" });
  });

  it("does not classify a valid primitive XML envelope as a workbook plan", async () => {
    const input = primitiveXml("read_range", { sheet: "Sheet1", address: "A1" });
    expect(isWorkbookPlanProtocolCandidate(input)).toBe(false);
    await expect(normalizeWorkbookPlanResponse(input)).resolves.toEqual({ ok: true, kind: "none" });
  });

  it("keeps primitive XML primitive when an argument value mentions the plan tool", async () => {
    const input = primitiveXml("read_range", {
      sheet: "Sheet1",
      address: "A1",
      note: "The literal string excel_execute_plan is workbook content, not a tool name.",
    });
    expect(isWorkbookPlanProtocolCandidate(input)).toBe(false);
    await expect(normalizeWorkbookPlanResponse(input)).resolves.toEqual({ ok: true, kind: "none" });
  });

  it("preclassifies malformed, multiple, and mixed XML containing a plan tool name", async () => {
    const plan = JSON.stringify({ name: "excel_execute_plan", arguments: validPlan() });
    const cases: ProtocolInput[] = [
      { content: `<tool_call>${plan}`, toolCalls: JSON.stringify([]) },
      { content: `<tool_call>${plan}</tool_call><tool_call>${plan}</tool_call>`, toolCalls: JSON.stringify([]) },
      {
        content: `<tool_call>${plan}</tool_call>`,
        toolCalls: JSON.stringify([{ id: "call_primitive", type: "function", function: { name: "read_range", arguments: "{}" } }]),
      },
    ];
    for (const input of cases) {
      expect(isWorkbookPlanProtocolCandidate(input)).toBe(true);
      await expectRejected(input);
    }
  });

  it("rejects plain JSON, including JSON embedded in hedging prose", async () => {
    await expectRejected(
      openAi([], `I think this is the plan: ${JSON.stringify(validPlan())}`),
    );
    await expectRejected(openAi([], JSON.stringify(validPlan())));
  });

  it("rejects success/completion prose attached to a plan before execution", async () => {
    await expectRejected(openAi([openAiCall(validPlan())], "Done — the workbook was updated successfully."));
    await expectRejected(xml(validPlan(), "Completed successfully. "));
  });

  it("rejects wrappers, routing/session fields, invented fields, and unsupported versions", async () => {
    await expectRejected(openAi([openAiCall({ plan: validPlan() })]));
    await expectRejected(openAi([openAiCall({ ...validPlan(), sessionId: "session" })]));
    await expectRejected(openAi([openAiCall({ ...validPlan(), policy: { allow: true } })]));
    await expectRejected(openAi([openAiCall({ ...validPlan(), invented: true })]));
    await expectRejected(openAi([openAiCall({ ...validPlan(), version: 2 })]));
  });

  it("rejects malformed and truncated JSON and XML boundaries", async () => {
    await expectRejected(
      openAi([
        {
          id: "call_plan_1",
          type: "function",
          function: { name: "excel_execute_plan", arguments: '{"version":1' },
        },
      ]),
    );
    await expectRejected(
      openAi([], '<tool_call>{"name":"excel_execute_plan","arguments":{"version":1}</tool_call>'),
    );
    const truncatedXml = xml(validPlan());
    await expectRejected({
      content: String(truncatedXml.content).replace("</tool_call>", ""),
      toolCalls: JSON.stringify([]),
    });
    await expectRejected(openAi([], "<tool_call>not-json</tool_call>"));
  });

  it("rejects wrong tool types, names, argument types, and envelope extras", async () => {
    await expectRejected(openAi([openAiCall(validPlan(), { type: "custom" })]));
    await expectRejected(openAi([openAiCall(validPlan(), { function: { name: "wrong", arguments: JSON.stringify(validPlan()) } })]));
    await expectRejected(openAi([openAiCall(validPlan(), { function: { name: "excel_execute_plan", arguments: validPlan() } })]));
    await expectRejected(openAi([openAiCall(validPlan(), { extra: true })]));
    await expect(
      normalizeWorkbookPlanResponse(
        openAi([], `<tool_call>${JSON.stringify({ name: "other_tool", arguments: validPlan() })}</tool_call>`),
      ),
    ).resolves.toEqual({ ok: true, kind: "none" });
    await expectRejected(
      openAi([], `<tool_call>${JSON.stringify({ name: "excel_execute_plan", arguments: JSON.stringify(validPlan()) })}</tool_call>`),
    );
  });

  it("rejects multiple, mixed, and cross-transport plan calls", async () => {
    await expectRejected(openAi([openAiCall(validPlan()), openAiCall(validPlan(), { id: "call_plan_2" })]));
    await expectRejected(
      openAi([
        openAiCall(validPlan()),
        { id: "call_primitive", type: "function", function: { name: "read_range", arguments: "{}" } },
      ]),
    );
    const xmlPlan = `<tool_call>${JSON.stringify({ name: "excel_execute_plan", arguments: validPlan() })}</tool_call>`;
    await expectRejected({ content: xmlPlan, toolCalls: JSON.stringify([openAiCall(validPlan())]) });
    await expectRejected({
      content: `${xmlPlan}<tool_call>${JSON.stringify({ name: "read_range", arguments: {} })}</tool_call>`,
      toolCalls: JSON.stringify([]),
    });
  });

  it("rejects ambiguous XML framing, extra tags, and primitive calls mixed with a plan", async () => {
    const plan = JSON.stringify({ name: "excel_execute_plan", arguments: validPlan() });
    await expectRejected({ content: `<tool_call>${plan}</tool_call> trailing`, toolCalls: JSON.stringify([]) });
    await expectRejected({ content: `<tool_call>${plan}</tool_call><other>text</other>`, toolCalls: JSON.stringify([]) });
    await expectRejected({ content: `<tool_call>${plan}</tool_call><tool_call>${plan}</tool_call>`, toolCalls: JSON.stringify([]) });
    await expectRejected({
      content: `<tool_call>${JSON.stringify({ name: "excel_execute_plan", arguments: validPlan() })}</tool_call>`,
      toolCalls: JSON.stringify([
        { id: "call_primitive", type: "function", function: { name: "read_range", arguments: "{}" } },
      ]),
    });
  });

  it("rejects accessor and throwing-proxy tool-call transports before any getter or trap", async () => {
    let accessed = 0;
    const accessorValue = {};
    Object.defineProperty(accessorValue, "arguments", {
      enumerable: true,
      get() {
        accessed += 1;
        throw new Error("getter canary");
      },
    });
    await expectRejected({ content: "", toolCalls: accessorValue as unknown as string });
    expect(accessed).toBe(0);

    let trapped = 0;
    const throwingProxy = new Proxy({}, {
      get() {
        trapped += 1;
        throw new Error("proxy canary");
      },
      ownKeys() {
        trapped += 1;
        throw new Error("proxy canary");
      },
      getOwnPropertyDescriptor() {
        trapped += 1;
        throw new Error("proxy canary");
      },
    });
    await expectRejected({ content: "", toolCalls: throwingProxy as unknown as string });
    expect(trapped).toBe(0);
  });

  it("rejects former outer, array, call, and function proxy positions content-free", async () => {
    let traps = 0;
    const hostileHandler: ProxyHandler<object> = {
      get() {
        traps += 1;
        throw new Error("get trap");
      },
      ownKeys() {
        traps += 1;
        throw new Error("ownKeys trap");
      },
      getOwnPropertyDescriptor() {
        traps += 1;
        throw new Error("descriptor trap");
      },
      getPrototypeOf() {
        traps += 1;
        throw new Error("prototype trap");
      },
    };
    const formerPositions: unknown[] = [
      new Proxy([openAiCall(validPlan())], hostileHandler),
      new Proxy({ 0: openAiCall(validPlan()), length: 1 }, hostileHandler),
      new Proxy(openAiCall(validPlan()), hostileHandler),
      new Proxy(function hostileFunction() {}, hostileHandler),
    ];
    for (const toolCalls of formerPositions) {
      const result = await normalizeWorkbookPlanResponse({
        content: "",
        toolCalls: toolCalls as string,
      });
      expect(isWorkbookPlanProtocolCandidate({ content: "", toolCalls: toolCalls as string })).toBe(true);
      expect(result).toEqual({
        ok: false,
        error: {
          code: "protocol_rejected",
          message: "Invalid workbook plan protocol envelope",
        },
      });
      expect(JSON.stringify(result)).not.toMatch(/Sheet1|A1|Write one|secret/i);
    }
    expect(traps).toBe(0);
  });

  it("does not turn benign quoted version/outcome prose into a protocol error", async () => {
    const finalAnswer = 'The note quoted "version" and "outcome" while describing the workbook.';
    const run = new HermesAgent();
    const excel = {
      run: vi.fn(async (work: (context: any) => Promise<unknown>) =>
        work({
          workbook: {
            worksheets: {
              getActiveWorksheet: () => ({ load: vi.fn(), name: "Sheet1" }),
              load: vi.fn(),
              items: [{ name: "Sheet1", getUsedRangeOrNullObject: () => ({ load: vi.fn(), isNullObject: true }) }],
            },
            getSelectedRange: () => ({ load: vi.fn(), address: "Sheet1!A1" }),
          },
          sync: vi.fn(),
        }),
      ),
    };
    vi.stubGlobal("Excel", excel);
    vi.stubGlobal("fetch", vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      headers: new Headers({ "content-type": "application/json" }),
      body: new ReadableStream({
        start(controller) {
          controller.enqueue(new TextEncoder().encode(JSON.stringify({
            choices: [{ message: { content: finalAnswer, tool_calls: [] } }],
          })));
          controller.close();
        },
      }),
    }));

    const callbacks = {
      onAssistantStart: vi.fn(),
      onAssistantDelta: vi.fn(),
      onAssistantFinal: vi.fn(),
      onActivity: vi.fn(),
      onStatus: vi.fn(),
    };
    const result = await run.run(
      "Explain the workbook note.",
      { baseUrl: "http://example.test/v1", model: "test", apiKey: "test", maxIterations: 1 },
      callbacks,
      new AbortController().signal,
    );

    expect(result.finalContent).toBe(finalAnswer);
    expect(callbacks.onAssistantFinal).toHaveBeenCalledWith(finalAnswer);
  });

  it("enforces input, tool-count, XML, and argument bounds before plan parsing", async () => {
    await expectRejected(openAi([openAiCall(validPlan())], "x".repeat(MAX_PROTOCOL_CONTENT_BYTES + 1)));
    await expectRejected({ content: "", toolCalls: JSON.stringify(Array.from({ length: MAX_PROTOCOL_TOOL_CALLS + 1 }, () => openAiCall(validPlan()))) });
    await expectRejected(openAi([openAiCall(validPlan(), { function: { name: "excel_execute_plan", arguments: "x".repeat(MAX_PROTOCOL_ARGUMENT_BYTES + 1) } })]));
    await expectRejected({ content: "", toolCalls: "x".repeat(MAX_PROTOCOL_TOOL_CALL_BYTES + 1) });
    await expectRejected({ content: `<tool_call>${"x".repeat(MAX_PROTOCOL_CONTENT_BYTES + 1)}</tool_call>`, toolCalls: JSON.stringify([]) });
  });

  it("accepts a previously bounded primitive structured call set above the old aggregate transport limit", async () => {
    const primitiveArgument = JSON.stringify({ value: "x".repeat(60 * 1024) });
    const serializedToolCalls = JSON.stringify(
      Array.from({ length: 3 }, (_, index) => ({
        id: `call_primitive_${index}`,
        type: "function",
        function: { name: "read_range", arguments: primitiveArgument },
      })),
    );
    expect(serializedToolCalls.length).toBeGreaterThan(128 * 1024);
    await expect(normalizeWorkbookPlanResponse({ content: null, toolCalls: serializedToolCalls })).resolves.toEqual({
      ok: true,
      kind: "none",
    });
  });
});
