import { beforeEach, describe, expect, it, vi } from "vitest";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

import { HermesAgent } from "../agent";
import { buildServerOwnedRequest } from "../serverOwnedMode";
import * as wordTools from "../wordTools";

const SERVER_MODEL = "hermes-agent";
const DIRECT_MODEL = "gpt-4.1-mini";
const BRIDGE_SESSION_ID = "660e8400-e29b-41d4-a716-446655440001";

const snapshot = {
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

function makeFinalResponse(content: string): Response {
  const body = new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode(JSON.stringify({ choices: [{ message: { content } }] })));
      controller.close();
    }
  });
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    body,
  } as unknown as Response;
}

function makePartialResponse(content: string): Response {
  const body = new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode(JSON.stringify({
        choices: [{
          message: { content },
          finish_reason: "length",
        }],
        hermes: {
          completed: false,
          partial: true,
          failed: false,
          error: "output truncated",
        },
      })));
      controller.close();
    },
  });
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    body,
  } as unknown as Response;
}

function makePartialStreamResponse(content: string): Response {
  const body = new ReadableStream({
    start(controller) {
      const events = [
        `data: ${JSON.stringify({ choices: [{ delta: { content } }] })}\n\n`,
        `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: "length" }], hermes: { completed: false, partial: true, failed: false } })}\n\n`,
        "data: [DONE]\n\n",
      ];
      for (const event of events) {
        controller.enqueue(new TextEncoder().encode(event));
      }
      controller.close();
    },
  });
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "text/event-stream" }),
    body,
  } as unknown as Response;
}

function makeToolCallResponse(toolName: string): Response {
  const body = new ReadableStream({
    start(controller) {
      controller.enqueue(new TextEncoder().encode(JSON.stringify({
        choices: [{
          message: {
            content: "iteration 1",
            tool_calls: [{
              id: "tool-1",
              type: "function",
              function: { name: toolName, arguments: "{}" }
            }]
          }
        }]
      })));
      controller.close();
    }
  });
  return {
    ok: true,
    status: 200,
    headers: new Headers({ "content-type": "application/json" }),
    body,
  } as unknown as Response;
}

const callbacks = {
  onAssistantStart: () => undefined,
  onAssistantDelta: () => undefined,
  onAssistantFinal: () => undefined,
  onActivity: () => undefined,
  onStatus: () => undefined,
};

describe("HermesAgent server-owned mode", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(snapshot as never);
    vi.spyOn(wordTools, "executeWordToolParsed").mockResolvedValue('{"ok":true}');
  });

  it("binds the exact bridge session UUID into the trusted header and omits caller tools", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("Server handled this natively.") as unknown as Response
    );

    const result = await new HermesAgent().run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        maxIterations: 4
      },
      callbacks,
      new AbortController().signal,
      BRIDGE_SESSION_ID,
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    const requestInit = fetchSpy.mock.calls[0]?.[1] as RequestInit | undefined;
    const headers = requestInit?.headers as Record<string, string> | undefined;
    expect(headers?.["X-Hermes-Session-Id"]).toBe(BRIDGE_SESSION_ID);
    const body = JSON.parse(String(requestInit?.body ?? "{}")) as {
      tools?: unknown;
      tool_choice?: unknown;
      messages?: Array<{ role: string; content?: string | null }>;
    };
    expect(body.tools).toBeUndefined();
    expect(body.tool_choice).toBeUndefined();
    expect(JSON.stringify(body.messages ?? [])).not.toContain("bridgeSessionId");
    expect(JSON.stringify(body.messages ?? [])).not.toContain(BRIDGE_SESSION_ID);
    expect(result.finalContent).toBe("Server handled this natively.");
  });

  it("fails closed when model=hermes-agent and the bridge session is missing", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("should not reach") as unknown as Response
    );

    await expect(new HermesAgent().run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        maxIterations: 4
      },
      callbacks,
      new AbortController().signal,
    )).rejects.toThrow(/bridge.*session/i);

    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("fails closed before any request when the bridge session ID is malformed", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("should not reach") as unknown as Response
    );

    await expect(new HermesAgent().run(
      "Summarize this document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        maxIterations: 4
      },
      callbacks,
      new AbortController().signal,
      "not-a-uuid",
    )).rejects.toThrow(/UUID/i);

    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it("names the current product, not the old Hermes for Word name, in both bridge-session errors", async () => {
    vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("should not reach") as unknown as Response
    );
    const settings = { baseUrl: "http://127.0.0.1:8645/v1", model: SERVER_MODEL, maxIterations: 4 };

    const agentError = await new HermesAgent().run(
      "Summarize this document", settings, callbacks, new AbortController().signal, "not-a-uuid",
    ).then(() => undefined, (error: unknown) => error);
    let requestError: unknown;
    try {
      buildServerOwnedRequest(settings, [], "not-a-uuid");
    } catch (error) {
      requestError = error;
    }

    for (const error of [agentError, requestError]) {
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).message).toContain("Open the Sidecar for Word task pane and try again.");
      expect((error as Error).message).not.toContain("Hermes for Word");
    }
  });

  it("ships the corrected bridge-session message in the plugin-served sidebar bundle", () => {
    const sidebar = "plugins/atlas-sidecar-word/sidebar";
    const bundles = readdirSync(sidebar).filter((name) => /^taskpane\.[0-9a-f]+\.js$/.test(name));
    expect(bundles).toHaveLength(1);
    const bundle = readFileSync(join(sidebar, bundles[0]), "utf8");
    expect(bundle).toContain("Open the Sidecar for Word task pane and try again.");
    expect(bundle).not.toContain("Hermes for Word");
  });

  it("allows the server-owned Hermes loop to exceed the direct-provider 20-second call budget", async () => {
    vi.useFakeTimers();
    try {
      vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 25_000));
        return makeFinalResponse("Completed after native tools.") as unknown as Response;
      });

      const run = new HermesAgent().run(
        "Read, edit, and verify this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      );

      await vi.advanceTimersByTimeAsync(25_000);
      await expect(run).resolves.toMatchObject({
        finalContent: "Completed after native tools.",
      });
    } finally {
      vi.useRealTimers();
    }
  });

  it("uses the full overview character count for one snapshot and permits the selected large budget", async () => {
    vi.useFakeTimers();
    try {
      const largeSnapshot = {
        ...snapshot,
        overview: { ...snapshot.overview, charCount: 50_000 },
        documentText: {
          ...snapshot.documentText,
          text: "truncated context",
          truncated: true,
          originalLength: 50_000,
          returnedLength: 17,
        },
      };
      const snapshotSpy = vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(largeSnapshot as never);
      const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 121_000));
        return makeFinalResponse("Completed within the selected large budget.") as unknown as Response;
      });

      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      );

      const outcome = run.then(
        (value) => ({ status: "fulfilled" as const, value }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );
      await vi.advanceTimersByTimeAsync(121_000);
      const settled = await outcome;
      expect(settled.status).toBe("fulfilled");
      if (settled.status === "fulfilled") {
        expect(settled.value.finalContent).toBe("Completed within the selected large budget.");
      }
      expect(snapshotSpy).toHaveBeenCalledTimes(1);
      expect(fetchSpy).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it("does not start model work after the caller aborts while the snapshot is handed off", async () => {
    const controller = new AbortController();
    const reason = new DOMException("Stopped by the user.", "AbortError");
    const snapshotWithAbort = {
      get overview() {
        controller.abort(reason);
        return snapshot.overview;
      },
      documentText: snapshot.documentText,
    };
    vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue(snapshotWithAbort as never);
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeFinalResponse("should not reach") as unknown as Response,
    );

    const outcome = new HermesAgent().run(
      "Proofread this document.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        ["api" + "Key"]: "",
        maxIterations: 4,
      } as never,
      callbacks,
      controller.signal,
      BRIDGE_SESSION_ID,
    ).then(
      () => ({ status: "fulfilled" as const }),
      (error: unknown) => ({ status: "rejected" as const, error }),
    );

    const settled = await outcome;
    expect(settled.status).toBe("rejected");
    if (settled.status === "rejected") {
      expect(settled.error).toBe(reason);
    }
    expect(fetchSpy).not.toHaveBeenCalled();
  });

  it.each<[number, number]>([
    [5_000, 120],
    [5_001, 150],
    [10_001, 180],
    [15_001, 210],
  ])("reports the selected server-owned timeout warning for %s characters", async (charCount, seconds) => {
    vi.useFakeTimers();
    try {
      vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue({
        ...snapshot,
        overview: { ...snapshot.overview, charCount },
      } as never);
      vi.spyOn(globalThis, "fetch").mockImplementation(() => new Promise<Response>(() => undefined));

      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(0);
      await vi.advanceTimersByTimeAsync(seconds * 1_000);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toMatchObject({
          name: "TimeoutError",
          message: expect.stringContaining(`${seconds}-second absolute time limit`),
        });
        expect((settled.error as Error).message).toContain(
          "Document actions may have occurred; review the document before retrying.",
        );
      }
    } finally {
      vi.useRealTimers();
    }
  });

  it("caps the largest server-owned document at 300 seconds", async () => {
    vi.useFakeTimers();
    try {
      vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue({
        ...snapshot,
        overview: { ...snapshot.overview, charCount: 1_000_000 },
      } as never);
      vi.spyOn(globalThis, "fetch").mockImplementation(() => new Promise<Response>(() => undefined));

      let settledBeforeMaximum = false;
      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => {
          settledBeforeMaximum = true;
          return { status: "fulfilled" as const };
        },
        (error: unknown) => {
          settledBeforeMaximum = true;
          return { status: "rejected" as const, error };
        },
      );

      await vi.advanceTimersByTimeAsync(0);
      await vi.advanceTimersByTimeAsync(299_999);
      expect(settledBeforeMaximum).toBe(false);
      await vi.advanceTimersByTimeAsync(1);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toMatchObject({
          name: "TimeoutError",
          message: expect.stringContaining("300-second absolute time limit"),
        });
      }
    } finally {
      vi.useRealTimers();
    }
  });

  it("charges snapshot latency against the selected absolute server-owned budget", async () => {
    vi.useFakeTimers();
    try {
      const largeSnapshot = {
        ...snapshot,
        overview: { ...snapshot.overview, charCount: 20_000 },
      };
      vi.spyOn(wordTools, "getDocumentSnapshot").mockImplementation(() => new Promise((resolve) => {
        setTimeout(() => resolve(largeSnapshot as never), 30_000);
      }));
      vi.spyOn(globalThis, "fetch").mockImplementation(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 195_000));
        return makeFinalResponse("too late") as unknown as Response;
      });

      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(30_000);
      await vi.advanceTimersByTimeAsync(180_000);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toMatchObject({
          name: "TimeoutError",
          message: expect.stringContaining("210-second absolute time limit"),
        });
      }
    } finally {
      vi.useRealTimers();
    }
  });

  it("bounds a hung server-owned snapshot with the 120-second startup guard and never fetches", async () => {
    vi.useFakeTimers();
    try {
      vi.spyOn(wordTools, "getDocumentSnapshot").mockReturnValue(new Promise<typeof snapshot>(() => undefined));
      const fetchSpy = vi.spyOn(globalThis, "fetch");
      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(0);
      await vi.advanceTimersByTimeAsync(119_999);
      expect(fetchSpy).not.toHaveBeenCalled();
      await vi.advanceTimersByTimeAsync(1);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toMatchObject({
          name: "TimeoutError",
          message: expect.stringContaining("120-second startup guard"),
        });
      }
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });

  it("does not start model work when a timed-out snapshot resolves late", async () => {
    vi.useFakeTimers();
    try {
      let resolveSnapshot!: (value: typeof snapshot) => void;
      vi.spyOn(wordTools, "getDocumentSnapshot").mockReturnValue(new Promise<typeof snapshot>((resolve) => {
        resolveSnapshot = resolve;
      }));
      const fetchSpy = vi.spyOn(globalThis, "fetch");
      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(120_000);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      resolveSnapshot(snapshot);
      await Promise.resolve();
      await Promise.resolve();
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });

  it("preserves the caller AbortError while the server-owned snapshot is pending", async () => {
    vi.useFakeTimers();
    try {
      const controller = new AbortController();
      const reason = new DOMException("Stopped by the user.", "AbortError");
      vi.spyOn(wordTools, "getDocumentSnapshot").mockReturnValue(new Promise<typeof snapshot>(() => undefined));
      const fetchSpy = vi.spyOn(globalThis, "fetch");
      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        controller.signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(0);
      controller.abort(reason);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toBe(reason);
      }
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });

  it("preserves the caller AbortError while server-owned model work is pending", async () => {
    vi.useFakeTimers();
    try {
      const controller = new AbortController();
      const reason = new DOMException("Stopped by the user.", "AbortError");
      const fetchSpy = vi.spyOn(globalThis, "fetch").mockImplementation(() => new Promise<Response>(() => undefined));
      const run = new HermesAgent().run(
        "Proofread this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        controller.signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(0);
      expect(fetchSpy).toHaveBeenCalledTimes(1);
      controller.abort(reason);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toBe(reason);
      }
    } finally {
      vi.useRealTimers();
    }
  });

  it("cleans the startup and selected-budget timers after server-owned completion", async () => {
    vi.useFakeTimers();
    try {
      vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue({
        ...snapshot,
        overview: { ...snapshot.overview, charCount: 50_000 },
      } as never);
      vi.spyOn(globalThis, "fetch").mockResolvedValue(
        makeFinalResponse("completed") as unknown as Response,
      );

      await new HermesAgent().run(
        "Summarize this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacks,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      );

      expect(vi.getTimerCount()).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it("does not start model work if the selected budget expires before the fetch begins", async () => {
    vi.useFakeTimers();
    try {
      vi.spyOn(wordTools, "getDocumentSnapshot").mockResolvedValue({
        ...snapshot,
        overview: { ...snapshot.overview, charCount: 20_000 },
      } as never);
      const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
        makeFinalResponse("should not reach") as unknown as Response,
      );
      const callbacksAtDeadline = {
        ...callbacks,
        onAssistantStart: () => {
          vi.advanceTimersByTime(210_000);
        },
      };

      const run = new HermesAgent().run(
        "Summarize this document.",
        {
          baseUrl: "http://127.0.0.1:8645/v1",
          model: SERVER_MODEL,
          ["api" + "Key"]: "",
          maxIterations: 4,
        } as never,
        callbacksAtDeadline,
        new AbortController().signal,
        BRIDGE_SESSION_ID,
      ).then(
        () => ({ status: "fulfilled" as const }),
        (error: unknown) => ({ status: "rejected" as const, error }),
      );

      await vi.advanceTimersByTimeAsync(0);
      const settled = await run;
      expect(settled.status).toBe("rejected");
      if (settled.status === "rejected") {
        expect(settled.error).toMatchObject({
          name: "TimeoutError",
          message: expect.stringContaining("210-second absolute time limit"),
        });
      }
      expect(fetchSpy).not.toHaveBeenCalled();
    } finally {
      vi.useRealTimers();
    }
  });

  it("keeps direct-provider mode unchanged with caller tools and local Office.js execution", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValue(
      makeToolCallResponse("replace_selection") as unknown as Response
    );

    const result = await new HermesAgent().run(
      "Rewrite this selected phrase.",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: DIRECT_MODEL,
        maxIterations: 1
      },
      callbacks,
      new AbortController().signal,
      undefined,
    );

    expect(fetchSpy).toHaveBeenCalledTimes(1);
    const requestInit = fetchSpy.mock.calls[0]?.[1] as RequestInit | undefined;
    const body = JSON.parse(String(requestInit?.body ?? "{}")) as {
      tools?: Array<{ function?: { name?: string } }>;
      tool_choice?: string;
    };
    expect(Array.isArray(body.tools)).toBe(true);
    expect((body.tools ?? []).length).toBeGreaterThan(0);
    expect(body.tool_choice).toBe("auto");
    expect(result.finalContent).toContain("I reached the max of 1 agent iterations");
  });

  it("rejects an empty server-owned final response instead of reporting Ready", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(makeFinalResponse(""));
    const statuses: string[] = [];
    const assistantFinals: string[] = [];

    await expect(new HermesAgent().run(
      "proof read the document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        maxIterations: 4,
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: (content) => assistantFinals.push(content),
        onActivity: () => undefined,
        onStatus: (status) => statuses.push(status),
      },
      new AbortController().signal,
      BRIDGE_SESSION_ID,
    )).rejects.toThrow("Hermes returned no usable final response");

    expect(assistantFinals).toEqual([]);
    expect(statuses).not.toContain("Ready");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("rejects a partial JSON server-owned response instead of reporting Ready", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(makePartialResponse("partial answer"));
    const statuses: string[] = [];
    const assistantFinals: string[] = [];
    const agent = new HermesAgent();

    await expect(agent.run(
      "proof read the document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        maxIterations: 4,
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: (content) => assistantFinals.push(content),
        onActivity: () => undefined,
        onStatus: (status) => statuses.push(status),
      },
      new AbortController().signal,
      BRIDGE_SESSION_ID,
    )).rejects.toThrow("Hermes returned an incomplete response");

    expect(assistantFinals).toEqual([]);
    expect(statuses).not.toContain("Ready");
    expect(agent.history.at(-1)?.role).not.toBe("assistant");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it("rejects a partial SSE server-owned response instead of reporting Ready", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch").mockResolvedValueOnce(makePartialStreamResponse("partial answer"));
    const statuses: string[] = [];
    const assistantFinals: string[] = [];
    const agent = new HermesAgent();

    await expect(agent.run(
      "proof read the document",
      {
        baseUrl: "http://127.0.0.1:8645/v1",
        model: SERVER_MODEL,
        maxIterations: 4,
      },
      {
        onAssistantStart: () => undefined,
        onAssistantDelta: () => undefined,
        onAssistantFinal: (content) => assistantFinals.push(content),
        onActivity: () => undefined,
        onStatus: (status) => statuses.push(status),
      },
      new AbortController().signal,
      BRIDGE_SESSION_ID,
    )).rejects.toThrow("Hermes returned an incomplete response");

    expect(assistantFinals).toEqual([]);
    expect(statuses).not.toContain("Ready");
    expect(agent.history.at(-1)?.role).not.toBe("assistant");
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });
});
