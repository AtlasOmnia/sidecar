// SPDX-License-Identifier: Apache-2.0
import { describe, expect, it, vi } from "vitest";
import {
  ApprovalRequiredError,
  DESK_APP,
  DESK_COMMAND,
  DESK_PORTS,
  DESK_SERVICE,
  DeskConnector,
  GRANT_HEADER,
} from "../deskConnect";
import { OfficeCompanionBridge } from "../officeBridge";

const WORKBENCH_PORTS = [47615, 47616, 47617, 47618, 47619];

class MemoryStorage {
  data = new Map<string, string>();
  getItem(key: string): string | null { return this.data.get(key) ?? null; }
  setItem(key: string, value: string): void { this.data.set(key, value); }
  removeItem(key: string): void { this.data.delete(key); }
}

function json(status: number, body: unknown): Response {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

describe("Sidecar connector", () => {
  it("uses its own loopback port range, never Workbench's range", () => {
    expect(DESK_PORTS).toHaveLength(5);
    expect(DESK_PORTS.some((p) => WORKBENCH_PORTS.includes(p))).toBe(false);
    expect(DESK_SERVICE).toBe(`hermes-desk-${DESK_APP}`);
    expect(DESK_COMMAND).toBe("/sidecar-excel");
    expect(new ApprovalRequiredError().message).toContain("/sidecar-excel");
  });

  it("discovers the first port whose hello names this app's bridge", async () => {
    const seen: string[] = [];
    const fetchImpl = vi.fn(async (url: string) => {
      seen.push(url);
      if (url.includes(`:${DESK_PORTS[0]}/`)) throw new TypeError("connection refused");
      if (url.includes(`:${DESK_PORTS[1]}/`)) return json(200, { service: "something-else", app: DESK_APP });
      return json(200, { service: DESK_SERVICE, app: DESK_APP });
    }) as unknown as typeof fetch;
    const connector = new DeskConnector({ fetchImpl, storage: new MemoryStorage() });
    expect(await connector.discover()).toBe(`https://localhost:${DESK_PORTS[2]}`);
    expect(seen[0]).toBe(`https://localhost:${DESK_PORTS[0]}/${DESK_APP}/hello`);
    expect(seen.every((url) => url.startsWith("https://localhost:"))).toBe(true);
    expect(connector.discoveryError).toBe("");
  });

  it("returns null when no bridge answers", async () => {
    const fetchImpl = vi.fn(async () => { throw new TypeError("refused"); }) as unknown as typeof fetch;
    expect(await new DeskConnector({ fetchImpl, storage: new MemoryStorage() }).discover()).toBeNull();
  });

  it("preserves the real network failure and explains ambiguous WebKit failures", async () => {
    const fetchImpl = vi.fn(async () => { throw new TypeError("Load failed"); }) as unknown as typeof fetch;
    const connector = new DeskConnector({ fetchImpl, storage: new MemoryStorage(), ports: [47640] });
    expect(await connector.discover()).toBeNull();
    expect(connector.discoveryError).toContain("https://localhost:47640: TypeError: Load failed");
    expect(connector.discoveryError).toContain("certificate");
    expect(connector.discoveryError).toContain("origin");
    expect(connector.discoveryError).not.toContain("is caused by");
  });

  it("distinguishes HTTP rejection, wrong service, timeout and connection refusal", async () => {
    const answers = [
      () => json(403, { error: "untrusted body must not be shown" }),
      () => json(200, { service: "another-app", app: "word" }),
      () => { throw new DOMException("request timed out", "AbortError"); },
      () => { throw new TypeError("connection refused"); },
    ];
    const fetchImpl = vi.fn(async () => answers.shift()!()) as unknown as typeof fetch;
    const connector = new DeskConnector({ fetchImpl, storage: new MemoryStorage(), ports: [47640, 47641, 47642, 47643] });
    expect(await connector.discover()).toBeNull();
    expect(connector.discoveryError).toContain("HTTP 403 (origin or host rejected)");
    expect(connector.discoveryError).toContain("not this app's bridge");
    expect(connector.discoveryError).toContain("AbortError: request timed out");
    expect(connector.discoveryError).toContain("TypeError: connection refused");
    expect(connector.discoveryError).not.toContain("untrusted body");
  });

  it("clears stale diagnostics on retry and bounds error text", async () => {
    let fails = true;
    const fetchImpl = vi.fn(async () => {
      if (fails) throw new Error("x".repeat(10000));
      return json(200, { service: DESK_SERVICE, app: DESK_APP });
    }) as unknown as typeof fetch;
    const connector = new DeskConnector({ fetchImpl, storage: new MemoryStorage(), ports: [47640] });
    await connector.discover();
    expect(connector.discoveryError.length).toBeLessThan(600);
    fails = false;
    expect(await connector.discover()).toBe("https://localhost:47640");
    expect(connector.discoveryError).toBe("");
  });

  it("waits for Allow in Hermes, then remembers the grant", async () => {
    const storage = new MemoryStorage();
    const polls = ["pending", "pending", "approved"];
    const bodies: unknown[] = [];
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      bodies.push(JSON.parse(String(init?.body ?? "null")));
      if (url.endsWith(`/${DESK_APP}/connect`)) {
        return json(201, { requestId: "r1", pollSecret: "s".repeat(43), code: "AB12-CD34", command: DESK_COMMAND });
      }
      const status = polls.shift();
      return json(200, status === "approved" ? { status, grant: "g".repeat(43) } : { status, code: "AB12-CD34" });
    }) as unknown as typeof fetch;
    const shown: string[] = [];
    const connector = new DeskConnector({ fetchImpl, storage, sleep: async () => undefined, client: { host: "Excel" } });
    const outcome = await connector.requestApproval("https://localhost:47640", (code, command) => shown.push(`${code} ${command}`));
    expect(outcome).toBe("approved");
    expect(shown).toEqual([`AB12-CD34 ${DESK_COMMAND}`]);
    expect(connector.grant).toBe("g".repeat(43));
    expect(bodies[0]).toEqual({ client: { host: "Excel" } });
    expect(bodies[1]).toEqual({ requestId: "r1", pollSecret: "s".repeat(43) });
    // A new page (Office restart) still has it.
    expect(new DeskConnector({ fetchImpl, storage }).grant).toBe("g".repeat(43));
    connector.clearGrant();
    expect(new DeskConnector({ fetchImpl, storage }).grant).toBeNull();
  });

  it("reports a denial without storing anything", async () => {
    const storage = new MemoryStorage();
    const fetchImpl = vi.fn(async (url: string) => url.endsWith("/connect")
      ? json(201, { requestId: "r1", pollSecret: "s".repeat(43), code: "AB12-CD34" })
      : json(200, { status: "denied" })) as unknown as typeof fetch;
    const connector = new DeskConnector({ fetchImpl, storage, sleep: async () => undefined });
    expect(await connector.requestApproval("https://localhost:47640", () => undefined)).toBe("denied");
    expect(connector.grant).toBeNull();
    expect(storage.data.size).toBe(0);
  });
});

describe("bridge session with a Sidecar grant", () => {
  const paneId = "11111111-1111-4111-8111-111111111111";
  function makeBridge(fetchImpl: typeof fetch, grant: string | null, onApprovalRequired?: () => void) {
    return new OfficeCompanionBridge({
      hostPrefix: DESK_APP,
      paneId,
      documentId: async () => "22222222-2222-4222-8222-222222222222",
      executeTool: async () => "{}",
      fetchImpl,
      baseUrl: "https://localhost:47640",
      grant: () => grant,
      onApprovalRequired,
    });
  }

  it("sends the grant only when registering, then exposes chat headers without any key", async () => {
    const calls: Array<{ url: string; headers: Record<string, string> }> = [];
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      calls.push({ url, headers: { ...(init?.headers as Record<string, string>) } });
      return json(201, { sessionId: "33333333-3333-4333-8333-333333333333", sessionKey: "k".repeat(64),
        protocolVersion: "1.0", expiresInSeconds: 60 });
    }) as unknown as typeof fetch;
    const bridge = makeBridge(fetchImpl, "g".repeat(43));
    expect(bridge.chatHeaders()).toBeNull();
    await bridge.connect();
    expect(calls[0].url).toBe(`https://localhost:47640/${DESK_APP}/session`);
    expect(calls[0].headers[GRANT_HEADER]).toBe("g".repeat(43));
    const headers = bridge.chatHeaders()!;
    expect(headers).toEqual({
      "X-Hermes-Session-Key": "k".repeat(64),
      "X-Hermes-Session-Id": "33333333-3333-4333-8333-333333333333",
      "X-Hermes-Pane-Id": paneId,
    });
    expect(JSON.stringify(headers)).not.toMatch(/Authorization|Grant/i);
  });

  it("turns Hermes' approval-required answer into ApprovalRequiredError", async () => {
    const onApprovalRequired = vi.fn();
    const fetchImpl = vi.fn(async () => json(403, { ok: false, approval: "required", error: "Hermes approval required" })) as unknown as typeof fetch;
    const bridge = makeBridge(fetchImpl, null, onApprovalRequired);
    await expect(bridge.connect()).rejects.toBeInstanceOf(ApprovalRequiredError);
    expect(onApprovalRequired).toHaveBeenCalledOnce();
  });
});
