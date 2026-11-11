import { afterEach, describe, expect, it, vi } from "vitest";
import {
  executeValidatedWordBridgeTool,
  getOrCreateWordDocumentId,
} from "../officeBridgeWord";

const DOCUMENT_ID = "eeeeeeee-eeee-4eee-8eee-eeeeeeeeeeee";

afterEach(() => {
  delete (globalThis as { Office?: unknown }).Office;
});

describe("Word bridge Office adapter", () => {
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
    expect(await getOrCreateWordDocumentId(() => DOCUMENT_ID)).toBe(DOCUMENT_ID);
    expect(values.get("hermes.officeBridge.documentId.v1")).toBe(DOCUMENT_ID);
    expect(settings.saveAsync).toHaveBeenCalledTimes(1);
    expect(await getOrCreateWordDocumentId(() => "ffffffff-ffff-4fff-8fff-ffffffffffff"))
      .toBe(DOCUMENT_ID);
    expect(settings.saveAsync).toHaveBeenCalledTimes(1);
  });

  it("returns a structured validation error before touching Office.js", async () => {
    const result = JSON.parse(await executeValidatedWordBridgeTool(
      "insert_text", {},
    ));
    expect(result.ok).toBe(false);
    expect(result.category).toBe("validation");
    expect(result.error).toContain("text");
  });
});
