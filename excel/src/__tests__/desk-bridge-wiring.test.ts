import { describe, expect, it, vi } from "vitest";
import { readFileSync } from "node:fs";
import { createExcelCompanionBridge } from "../officeBridgeExcel";
import { ApprovalRequiredError, GRANT_HEADER } from "../deskConnect";
import { WORKBOOK_PLAN_FINGERPRINT } from "../generated/workbookPlanSchema";

const sessionId = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa";
const paneId = "bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb";
const documentId = "cccccccc-cccc-4ccc-8ccc-cccccccccccc";

describe("Sidecar Excel production bridge wiring", () => {
  it("sends the remembered grant only at session creation and authenticates chat with pane session headers", async () => {
    const fetchImpl = vi.fn(async () => new Response(JSON.stringify({
      sessionId, sessionKey: "s".repeat(64), documentId, protocolVersion: "1.0",
      expiresInSeconds: 60, workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT,
    }), { status: 201 }));
    const bridge = createExcelCompanionBridge(undefined, async () => "{}", {
      baseUrl: "https://localhost:47640", grant: () => "g".repeat(43),
      documentId: async () => documentId, paneId, fetchImpl,
    });
    expect(bridge.chatHeaders()).toBeNull();
    await bridge.connect();
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("https://localhost:47640/excel/session");
    expect(new Headers(init.headers).get(GRANT_HEADER)).toBe("g".repeat(43));
    expect(JSON.parse(String(init.body)).workbookPlanFingerprint).toBe(WORKBOOK_PLAN_FINGERPRINT);
    expect(bridge.chatHeaders()).toEqual({
      "X-Hermes-Session-Id": sessionId, "X-Hermes-Session-Key": "s".repeat(64),
      "X-Hermes-Pane-Id": paneId,
    });
    expect(bridge.chatHeaders()).not.toHaveProperty(GRANT_HEADER);
  });

  it("returns to human approval on a rejected or revoked grant without starting a session", async () => {
    const onApprovalRequired = vi.fn();
    const bridge = createExcelCompanionBridge(undefined, async () => "{}", {
      grant: () => "revoked-grant", onApprovalRequired,
      documentId: async () => documentId, paneId,
      fetchImpl: vi.fn(async () => new Response(JSON.stringify({ approval: "required" }), { status: 403 })),
    });
    await expect(bridge.connect()).rejects.toBeInstanceOf(ApprovalRequiredError);
    expect(onApprovalRequired).toHaveBeenCalledOnce();
    expect(bridge.chatHeaders()).toBeNull();
  });

  it("wires discovery, the human connect button and authenticated chat without a profile or credential form", () => {
    const html = readFileSync("src/taskpane/taskpane.html", "utf8");
    const pane = readFileSync("src/taskpane/taskpane.ts", "utf8");
    for (const id of ["desk-connect", "desk-connect-button", "desk-connect-code"]) expect(html).toContain(`id="${id}"`);
    expect(html).not.toMatch(/omnia-profile-chip|api-key-input|base-url-input/);
    expect(pane).toContain("connector.discover()");
    expect(pane).toContain("connector.requestApproval(");
    expect(pane).toContain("requestHeaders: chatHeaders");
    expect(pane).toContain("serverOwnedPlanRoute.executeCompanionTool");
    expect(pane).not.toMatch(/OmniaProfileMonitor|\/v1\/omnia\/profile/);
  });
});
