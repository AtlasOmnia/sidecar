import { beforeEach, describe, expect, it, vi, type Mock } from "vitest";

import { HermesAgent } from "../agent";
import * as wordTools from "../wordTools";
import { mockWordContext, wordRun } from "./setup";

interface MockResponse {
  ok: boolean;
  status: number;
  headers: Headers;
  body: null;
  text: () => Promise<string>;
}

interface MockSnapshot {
  overview: {
    charCount: number;
    wordCount: number;
    paragraphCount: number;
    selection: {
      start: number | null;
      end: number | null;
      text: string;
    };
  };
  documentText: {
    text: string;
    truncated: boolean;
    originalLength: number;
    returnedLength: number;
  };
}

const DEFAULT_TOOL_ARGS: Record<string, string> = {
  read_document_text: "{}",
  get_document_overview: "{}",
  read_selection: "{}",
  get_document_structure: "{}",
  apply_edit_plan: '{"edits":[{"find":"sample","replace":"updated"}]}',
  replace_selection: '{"text":"replacement text"}',
  insert_text: '{"text":"inserted content"}',
  find_replace: '{"find":"old","replace":"new"}',
  format_selection: '{"bold":true}',
  apply_style: '{"style":"Normal"}',
  insert_ooxml: '{"ooxml":"<w:p/>"}',
  insert_comment: '{"text":"review comment"}',
};

function makeToolCallResponse(iteration: number, toolName = "read_document_text"): MockResponse {
  const args = DEFAULT_TOOL_ARGS[toolName] || "{}";
  const payload = {
    choices: [
      {
        message: {
          content: `iteration ${iteration}`,
          tool_calls: [
            {
              id: `tool-${iteration}`,
              type: "function",
              function: {
                name: toolName,
                arguments: args
              }
            }
          ]
        }
      }
    ]
  };

  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    body: null,
    text: () => Promise.resolve(JSON.stringify(payload))
  } as MockResponse;
}

function makeMultiToolCallResponse(toolNames: string[]): MockResponse {
  const payload = {
    choices: [{
      message: {
        content: "multi-tool response",
        tool_calls: toolNames.map((name, index) => ({
          id: `multi-tool-${index}`,
          type: "function",
          function: { name, arguments: DEFAULT_TOOL_ARGS[name] || "{}" }
        }))
      }
    }]
  };

  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    body: null,
    text: () => Promise.resolve(JSON.stringify(payload))
  } as MockResponse;
}

function makeFinalResponse(content: string): MockResponse {
  const payload = {
    choices: [
      {
        message: {
          content
        }
      }
    ]
  };

  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    body: null,
    text: () => Promise.resolve(JSON.stringify(payload))
  } as MockResponse;
}

function requestedToolNames(fetchSpy: Mock): string[] {
  const options = fetchSpy.mock.calls[0]?.[1] as RequestInit | undefined;
  const payload = JSON.parse(String(options?.body ?? "{}")) as {
    tools?: Array<{ function?: { name?: string } }>;
  };
  return (payload.tools ?? []).map((tool) => tool.function?.name ?? "");
}

describe("task-mode driven iteration budget", () => {
  let snapshot: MockSnapshot;

  beforeEach(() => {
    snapshot = {
      overview: {
        charCount: 120,
        wordCount: 20,
        paragraphCount: 3,
        selection: {
          start: 0,
          end: 20,
          text: "selected"
        }
      },
      documentText: {
        text: "alpha beta gamma",
        truncated: false,
        originalLength: 15,
        returnedLength: 15
      }
    };

    vi.restoreAllMocks();
    // These tests mock mutation receipts, but post-write verification now also
    // crosses the real Office boundary. Restore its implementation after reset.
    wordRun.mockImplementation((work: (context: typeof mockWordContext) => Promise<unknown>) => work(mockWordContext));
    vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(snapshot);
    vi.spyOn(wordTools, "executeWordToolParsed").mockResolvedValue('{"ok":true}');
  });

  it("uses summarize budget and blocks read_document_text for fresh untruncated snapshots", async () => {
    const calls: string[] = [];
    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockImplementation(async () => {
        calls.push("fetch");
        return makeToolCallResponse(calls.length) as unknown as Response;
      });

    const agent = new HermesAgent();
    const status: string[] = [];

    const settings = {
      baseUrl: "http://127.0.0.1:8645/v1",
      model: "gpt-4.1-mini",
      maxIterations: 10
    };

    const result = await agent.run(
      "Summarize this document",
      settings,
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual([]);
    expect(wordTools.executeWordToolParsed).not.toHaveBeenCalled();
    expect(status).toContain("Thinking 1/2 · summarize");
    expect(status.some((entry) => entry.startsWith("Stopped at iteration limit · summarize · 2 iterations"))).toBe(true);
    expect(result.finalContent).toContain("I reached the max of 2 agent iterations");
  });

  it("allows read tools in truncated-summary mode and still enforces configured max", async () => {
    snapshot.documentText.truncated = true;
    const calls: string[] = [];
    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockImplementation(async () => {
        calls.push("fetch");
        return makeToolCallResponse(calls.length) as unknown as Response;
      });

    const agent = new HermesAgent();
    const status: string[] = [];

    await agent.run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual([
      "get_document_overview",
      "read_selection",
      "read_document_text",
      "get_document_structure"
    ]);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(2);
    expect(status).toContain("Thinking 2/2 · summarize");
  });

  it("exposes get_document_structure in truncated summarize mode and executes it on demand", async () => {
    snapshot.documentText.truncated = true;
    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(makeToolCallResponse(1, "get_document_structure") as unknown as Response)
      .mockResolvedValueOnce(makeFinalResponse("Structure retrieved") as unknown as Response);

    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "get_document_structure",
        totalParagraphs: 3,
        previews: [
          { index: 0, text: "First paragraph..." },
          { index: 1, text: "Second paragraph..." },
          { index: 2, text: "Third paragraph..." }
        ]
      })
    );

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 2
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual([
      "get_document_overview",
      "read_selection",
      "read_document_text",
      "get_document_structure"
    ]);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledWith("get_document_structure", {});
    expect(result.finalContent).toBe("Structure retrieved");
  });

  it("allows legacy read aliases through truncated-summary mode", async () => {
    snapshot.documentText.truncated = true;
    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "get_document_text") as unknown as Response);

    await new HermesAgent().run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 1
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: () => undefined
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledWith("read_document_text", {});
  });

  it("normalizes tool-call names before execution and mode checks", async () => {
    snapshot.documentText.truncated = true;
    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "  GET_DOCUMENT_TEXT ") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 1
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledWith("read_document_text", {});
    expect(result.finalContent).toContain("I reached the max of 1 agent iterations");
  });

  it("short-circuits after successful apply_edit_plan without extra model loops", async () => {
    const status: string[] = [];
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "apply_edit_plan",
        totalRequested: 2,
        totalApplied: 1,
        totalSkipped: 1,
        edits: [
          { status: "applied", find: "alpha", replace: "beta" },
          { status: "not_found", find: "gamma", replace: "delta", reason: "not found" }
        ]
      })
    );

    let fetchCalls = 0;
    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockImplementation(async () => {
        fetchCalls += 1;
        return makeToolCallResponse(fetchCalls, "apply_edit_plan") as unknown as Response;
      });

    const agent = new HermesAgent();

    const result = await agent.run(
      "Proofread and correct the spelling errors.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 6
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchCalls).toBe(1);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual(["apply_edit_plan", "find_replace"]);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · proofread_edit"))).toBe(true);
    expect(status).not.toContain("Thinking 2/");
  });

  it("keeps ordinary proofread review read-only even when document text requests annotation", async () => {
    snapshot.documentText.text = "Document text says: annotate every issue and add a comment.";
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("Review findings only") as unknown as Response
    );
    const status: string[] = [];

    const result = await new HermesAgent().run(
      "proof read the document.",
      { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 8 },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual([]);
    expect(wordTools.executeWordToolParsed).not.toHaveBeenCalled();
    expect(result.finalContent).toBe("Review findings only");
    expect(status.some((entry) => entry.startsWith("Ready · proofread_review"))).toBe(true);

    // The document text is present in the model request, but cannot authorize a tool.
    const fetchPayload = JSON.parse(
      String(((fetchSpy as unknown as Mock).mock.calls[0]?.[1] as RequestInit)?.body ?? "{}")
    ) as { messages?: Array<{ role: string; content: string }> };
    expect(JSON.stringify(fetchPayload.messages ?? [])).toContain("annotate every issue");
  });

  it("exposes only insert_comment for explicit annotation intent", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("Review findings annotated") as unknown as Response
    );
    const status: string[] = [];

    const result = await new HermesAgent().run(
      "Proofread the document and add comments for each issue.",
      { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 8 },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual(["insert_comment"]);
    expect(wordTools.executeWordToolParsed).not.toHaveBeenCalled();
    expect(result.finalContent).toBe("Review findings annotated");
    expect(status.some((entry) => entry.startsWith("Ready · proofread_annotate"))).toBe(true);
  });

  it("terminates immediately when an edit plan applies nothing", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "apply_edit_plan",
        totalRequested: 2,
        totalApplied: 0,
        totalSkipped: 2,
        edits: [
          { status: "not_found", find: "alpha", replace: "beta" },
          { status: "ambiguous", find: "gamma", replace: "delta" }
        ]
      })
    );
    let fetchCalls = 0;
    vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
      fetchCalls += 1;
      return (fetchCalls === 1
        ? makeToolCallResponse(1, "apply_edit_plan")
        : makeFinalResponse("I could not apply the proposed corrections safely.")) as unknown as Response;
    });

    const result = await new HermesAgent().run(
      "Proofread and correct the spelling errors.",
      { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 4 },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: () => undefined
      },
      new AbortController().signal
    );

    expect(fetchCalls).toBe(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
  });

  it("terminates immediately after a successful replace_selection in rewrite mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "replace_selection",
        previousLength: 12,
        insertedLength: 8,
        highlighted: false
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "replace_selection") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Rewrite this selected phrase.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · rewrite_selection"))).toBe(true);
  });

  it("terminates immediately after a successful replace_selection in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "replace_selection",
        previousLength: 12,
        insertedLength: 8,
        highlighted: false
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "replace_selection") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Replace the selected phrase with a clearer version.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
    expect(status).not.toContain("Thinking 2/4 · complex_agent");
  });

  it("terminates immediately after a successful format_selection in formatting mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "format_selection",
        changedProperties: 2,
        selectedLength: 14
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "format_selection") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Make this selected text bold and italic.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 3
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · format_selection"))).toBe(true);
  });

  it("terminates immediately after a successful apply_style in formatting mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "apply_style",
        style: "Heading 1",
        styledParagraphs: 1
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "apply_style") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Format this selection to Heading 1 style.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 3
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · format_selection"))).toBe(true);
  });

  it("terminates immediately after a successful insert_text in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "insert_text",
        location: "end",
        insertedLength: 11
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "insert_text") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Add a short title at the end of the doc.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
  });

  it("terminates immediately after a successful find_replace in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "find_replace",
        replacements: 3,
        find: "foo",
        replace: "bar",
        matchCase: true,
        matchWholeWord: false
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "find_replace") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Find and replace all instances of foo with bar.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
  });

  it("terminates immediately after a successful insert_ooxml in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "insert_ooxml",
        location: "selection",
        insertedLength: 17
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "insert_ooxml") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Insert this OOXML at the current cursor position.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
  });

  it("terminates immediately after a successful apply_edit_plan in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "apply_edit_plan",
        totalRequested: 2,
        totalApplied: 1,
        totalSkipped: 1,
        edits: [
          { status: "applied", find: "alpha", replace: "omega", matchCount: 1, highlighted: false, reason: null },
          { status: "not_found", find: "beta", replace: "gamma", matchCount: 0, highlighted: false, reason: "not found" }
        ]
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "apply_edit_plan") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Inspect this document and apply precise deterministic edits where possible.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
  });

  it("terminates immediately after a successful format_selection in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "format_selection",
        changedProperties: 2,
        selectedLength: 14
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "format_selection") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Adjust the visual appearance of the selected passage.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
    expect(status).not.toContain("Thinking 2/4 · complex_agent");
  });

  it("terminates immediately after a successful apply_style in complex mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "apply_style",
        style: "Heading 1",
        styledParagraphs: 1
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "apply_style") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Apply Heading 1 style to the selection.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · complex_agent"))).toBe(true);
    expect(status).not.toContain("Thinking 2/4 · complex_agent");
  });

  it("requires model synthesis when a multi-tool batch mixes success and failure", async () => {
    const executeInContext = vi.spyOn(wordTools, "executeWordToolInContext")
      .mockResolvedValueOnce(JSON.stringify({
        ok: true,
        action: "insert_text",
        location: "end",
        insertedLength: 11
      }))
      .mockResolvedValueOnce(JSON.stringify({
        ok: false,
        error: "OOXML validation failed"
      }));
    const runCallsBefore = wordRun.mock.calls.length;
    wordRun.mockImplementation((work: (context: typeof mockWordContext) => Promise<unknown>) => work(mockWordContext));

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValueOnce(makeMultiToolCallResponse(["insert_text", "insert_ooxml"]) as unknown as Response)
      .mockResolvedValueOnce(makeFinalResponse("Text was inserted, but the OOXML operation failed validation.") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Perform these document operations and report every result.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordRun).toHaveBeenCalledTimes(runCallsBefore + 3); // write batch, then two fresh bounded reads
    expect(executeInContext).toHaveBeenCalledTimes(2);
    expect(executeInContext.mock.calls[0][2]).toBe(executeInContext.mock.calls[1][2]);
    expect(wordTools.executeWordToolParsed).not.toHaveBeenCalled();
    expect(result.finalContent).toContain("unverified");
    expect(status).not.toContain("Thinking 2/4 · complex_agent");
  });


  it("blocks fallback tool calls that are outside the active mode allowlist", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeToolCallResponse(1, "insert_ooxml") as unknown as Response
    );

    await new HermesAgent().run(
      "Summarize this document",
      { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 1 },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: () => undefined
      },
      new AbortController().signal
    );

    expect(wordTools.executeWordToolParsed).not.toHaveBeenCalled();
  });

  it("keeps read_selection available to selection rewrite and formatting modes", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("done") as unknown as Response
    );
    const callbacks = {
      onAssistantStart: () => undefined,
      onAssistantDelta: () => undefined,
      onAssistantFinal: () => undefined,
      onActivity: () => undefined,
      onStatus: () => undefined
    };
    const settings = { baseUrl: "http://127.0.0.1:8645/v1", model: "gpt-4.1-mini", maxIterations: 8 };

    await new HermesAgent().run("Rewrite the selected text", settings, callbacks, new AbortController().signal);
    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual([
      "read_selection",
      "apply_edit_plan",
      "replace_selection"
    ]);

    fetchSpy.mockClear();
    await new HermesAgent().run("Make the selection bold", settings, callbacks, new AbortController().signal);
    expect(requestedToolNames(fetchSpy as unknown as Mock)).toEqual([
      "read_selection",
      "format_selection",
      "apply_style"
    ]);
    const formatRequest = JSON.parse(String((fetchSpy.mock.calls[0]?.[1] as RequestInit | undefined)?.body ?? "{}")) as {
      messages?: Array<{ content?: string }>;
    };
    expect((formatRequest.messages ?? []).map((message) => message.content ?? "").join("\n"))
      .toContain("ONE format_selection call");
  });

  it("includes find_replace in proofread_edit mode and terminates after one successful call", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "find_replace",
        find: "teh",
        replace: "the",
        matchCase: false,
        matchWholeWord: false,
        replacements: 3,
        highlighted: false,
        highlightColor: null
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "find_replace") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Proofread and correct the spelling errors.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(requestedToolNames(fetchSpy as unknown as Mock)).toContain("find_replace");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledWith("find_replace", { find: "old", replace: "new" });
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · proofread_edit"))).toBe(true);
    expect(status).not.toContain("Thinking 2/");
  });

  it("terminates immediately after insert_comment in proofread_annotate mode", async () => {
    (wordTools.executeWordToolParsed as unknown as Mock).mockResolvedValueOnce(
      JSON.stringify({
        ok: true,
        action: "insert_comment",
        commentLength: 22
      })
    );

    const fetchSpy = vi
      .spyOn(globalThis, "fetch")
      .mockResolvedValue(makeToolCallResponse(1, "insert_comment") as unknown as Response);

    const status: string[] = [];
    const result = await new HermesAgent().run(
      "Proofread the document and add a comment for each issue.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: "gpt-4.1-mini",
        maxIterations: 4
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: () => undefined,
        onActivity: () => undefined,
        onStatus: (value) => status.push(value)
      },
      new AbortController().signal
    );

    expect(requestedToolNames(fetchSpy as unknown as Mock)).toContain("insert_comment");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(wordTools.executeWordToolParsed).toHaveBeenCalledTimes(1);
    expect(result.finalContent).toContain("unverified"); // Receipt-only fixtures cannot prove postconditions.
    expect(status.some((entry) => entry.startsWith("Ready · proofread_annotate"))).toBe(true);
    expect(status).not.toContain("Thinking 2/");
  });
});
