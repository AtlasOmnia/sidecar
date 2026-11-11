import { afterEach, describe, expect, it, vi } from "vitest";
import {
  createExcelCompanionBridge,
  executeAcceptedWorkbookPlanThroughExcelBridge,
  executeValidatedExcelBridgeTool,
  getOrCreateExcelDocumentId,
} from "../officeBridgeExcel";
import { BrowserIdempotencyStore, createTrustedOperationBoundary, MemoryIdempotencyStorage } from "../workbookPlanExecutor";

const DOCUMENT_ID = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";

afterEach(() => {
  delete (globalThis as { Office?: unknown }).Office;
});

describe("Excel bridge Office adapter", () => {
  it("persists a document ID through Office document settings", async () => {
    const values = new Map<string, unknown>();
    const settings = {
      get: vi.fn((key: string) => values.get(key)),
      set: vi.fn((key: string, value: unknown) => values.set(key, value)),
      saveAsync: vi.fn((callback: (result: { status: string }) => void) =>
        callback({ status: "succeeded" })),
    };
    (globalThis as { Office?: unknown }).Office = {
      context: { document: { settings } },
      AsyncResultStatus: { Succeeded: "succeeded" },
    };
    expect(await getOrCreateExcelDocumentId(() => DOCUMENT_ID)).toBe(DOCUMENT_ID);
    expect(values.get("hermes.officeBridge.documentId.v1")).toBe(DOCUMENT_ID);
    expect(settings.saveAsync).toHaveBeenCalledTimes(1);
    expect(await getOrCreateExcelDocumentId(() => "ffffffff-ffff-4fff-8fff-ffffffffffff"))
      .toBe(DOCUMENT_ID);
    expect(settings.saveAsync).toHaveBeenCalledTimes(1);
  });

  it("returns a structured validation error before touching Office.js", async () => {
    const result = JSON.parse(await executeValidatedExcelBridgeTool(
      "read_range", { sheet: "Sheet1" },
    ));
    expect(result.ok).toBe(false);
    expect(result.category).toBe("validation");
    expect(result.error).toContain("address");
  });

  it("injects fetch, document id and pane id into createExcelCompanionBridge without Office", async () => {
    const paneId = "abababab-abab-4bab-8bab-abababababab";
    const fetchMock = vi.fn(async () => new Response(JSON.stringify({
      ok: true,
      sessionId: "cdcdcdcd-cdcd-4dcd-8dcd-cdcdcdcdcdcd",
      sessionKey: "k".repeat(64),
      protocolVersion: "1.0",
      expiresInSeconds: 60,
    }), { status: 201, headers: { "Content-Type": "application/json" } }));
    const create = createExcelCompanionBridge as unknown as (
      onState: undefined,
      executeTool: undefined,
      options: { fetchImpl: typeof fetch; documentId: () => Promise<string>; paneId: string },
    ) => ReturnType<typeof createExcelCompanionBridge>;
    const bridge = create(undefined, undefined, {
      fetchImpl: fetchMock as unknown as typeof fetch,
      documentId: async () => DOCUMENT_ID,
      paneId,
    });

    await expect(bridge.connect()).resolves.toMatchObject({ documentId: DOCUMENT_ID });
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(JSON.parse(String(init.body))).toMatchObject({ documentId: DOCUMENT_ID, paneId });
  });

  it("exposes a dependency-injected plan seam without changing primitive validation", async () => {
    let adapterCalls = 0;
    const report = await executeAcceptedWorkbookPlanThroughExcelBridge(
      Object.freeze({
        transactionId: DOCUMENT_ID,
        semanticDigest: "digest",
        operationCount: 0,
      }),
      createTrustedOperationBoundary({ sessionId: DOCUMENT_ID, documentId: DOCUMENT_ID, turnId: DOCUMENT_ID }),
      {
        executeFirstGroup: async () => {
          adapterCalls += 1;
          return { status: "committed" };
        },
        executeGroup: async () => {
          adapterCalls += 1;
          return { status: "committed" };
        },
        reconcileFailedGroup: async () => [],
      },
      new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    );
    expect(report.outcome).toBe("rejected");
    expect(adapterCalls).toBe(0);
  });
});
