// SPDX-License-Identifier: Apache-2.0
// Copyright 2026 Atlas Omnia
/**
 * Sidecar connection: find this app's HTTPS localhost bridge (bound to 127.0.0.1 by the plugin) and get the
 * user's approval in Hermes. The same file ships in the Word, Excel and PowerPoint sidebars; only the
 * DESK_APP constants differ.
 *
 * - The bridge answers only this sidebar's HTTPS origin, on its own app's port range.
 * - "Connect to Hermes" files a request; the user allows it in Hermes (/sidecar-<app>). The bridge then
 *   returns a per-sidebar grant ONCE; it is kept in this origin's localStorage so the approval survives
 *   Office and Hermes restarts. Hermes stores only a hash of it and can revoke it at any time.
 * - No API key ever reaches this page.
 */

export const DESK_APP = "word";
export const DESK_PORTS: readonly number[] = [47630, 47631, 47632, 47633, 47634];
export const DESK_SERVICE = `hermes-desk-${DESK_APP}`;
export const DESK_COMMAND = `/sidecar-${DESK_APP}`;
export const GRANT_HEADER = "X-Hermes-Desk-Grant";
const GRANT_STORAGE_KEY = `hermes.desk.${DESK_APP}.grant.v1`;
const POLL_INTERVAL_MS = 1500;
const HELLO_TIMEOUT_MS = 1500;

export type ApprovalOutcome = "approved" | "denied" | "expired";

export interface DeskStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export interface DeskConnectorOptions {
  fetchImpl?: typeof fetch;
  storage?: DeskStorage;
  sleep?: (ms: number) => Promise<void>;
  ports?: readonly number[];
  client?: Record<string, string>;
}

export class ApprovalRequiredError extends Error {
  constructor(message = `Hermes approval required. Click Connect to Hermes, then allow it in Hermes with ${DESK_COMMAND}.`) {
    super(message);
    this.name = "ApprovalRequiredError";
  }
}

export class DeskConnector {
  private readonly fetchImpl: typeof fetch;
  private readonly storage: DeskStorage | null;
  private readonly sleep: (ms: number) => Promise<void>;
  private readonly ports: readonly number[];
  private readonly client: Record<string, string>;
  private memoryGrant: string | null = null;
  private discoveryFailures: string[] = [];

  /** Browser errors are often opaque: report the observed error, not a guessed root cause. */
  get discoveryError(): string {
    if (!this.discoveryFailures.length) return "";
    return this.discoveryFailures.join("; ") +
      ". Check that the plugin is running and its local CA certificate is trusted. " +
      "A generic Load failed/Failed to fetch can mean certificate trust, connection refusal, origin rejection or a webview transport restriction; it does not identify which.";
  }

  constructor(options: DeskConnectorOptions = {}) {
    const fetchImpl = options.fetchImpl ?? globalThis.fetch;
    this.fetchImpl = fetchImpl.bind(globalThis);
    this.storage = options.storage ?? safeLocalStorage();
    this.sleep = options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.ports = options.ports ?? DESK_PORTS;
    this.client = options.client ?? {};
  }

  /** The first bridge on this app's port range that identifies as this app's Sidecar bridge. */
  async discover(): Promise<string | null> {
    this.discoveryFailures = [];
    for (const port of this.ports) {
      const base = `https://localhost:${port}`;
      const controller = typeof AbortController === "undefined" ? null : new AbortController();
      const timer = controller ? setTimeout(() => controller.abort(), HELLO_TIMEOUT_MS) : null;
      try {
        const response = await this.fetchImpl(`${base}/${DESK_APP}/hello`, {
          method: "GET",
          cache: "no-store",
          signal: controller?.signal,
        });
        if (!response.ok) {
          const hint = response.status === 403 || response.status === 421 ? " (origin or host rejected)" : "";
          this.discoveryFailures.push(`${base}: HTTP ${response.status}${hint}`);
          continue;
        }
        const body = (await response.json()) as { service?: unknown; app?: unknown };
        if (body && body.service === DESK_SERVICE && body.app === DESK_APP) {
          this.discoveryFailures = [];
          return base;
        }
        this.discoveryFailures.push(`${base}: not this app's bridge`);
      } catch (error) {
        const detail = error instanceof Error ? `${error.name}: ${error.message}` : "Unknown fetch error";
        this.discoveryFailures.push(`${base}: ${detail.replace(/[\r\n\u0000-\u001f]/g, " ").slice(0, 160)}`);
      } finally {
        if (timer) clearTimeout(timer);
      }
    }
    return null;
  }

  get grant(): string | null {
    try {
      const stored = this.storage?.getItem(GRANT_STORAGE_KEY) ?? null;
      return stored && stored.length >= 32 && stored.length <= 200 ? stored : this.memoryGrant;
    } catch {
      return this.memoryGrant;
    }
  }

  clearGrant(): void {
    this.memoryGrant = null;
    try {
      this.storage?.removeItem(GRANT_STORAGE_KEY);
    } catch {
      // Storage may be unavailable; the in-memory copy is already gone.
    }
  }

  private saveGrant(grant: string): void {
    this.memoryGrant = grant;
    try {
      this.storage?.setItem(GRANT_STORAGE_KEY, grant);
    } catch {
      // Without storage the approval lasts for this page only.
    }
  }

  /**
   * File a connect request and wait for the user's decision in Hermes. `onCode` shows the code the user
   * sees in Hermes' Allow/Deny question.
   */
  async requestApproval(
    base: string,
    onCode: (code: string, command: string) => void,
    signal?: AbortSignal,
  ): Promise<ApprovalOutcome> {
    const created = await this.post(base, `/${DESK_APP}/connect`, { client: this.client }, signal);
    if (created.status !== 201) {
      throw new Error(messageOf(created.body, "Hermes did not accept the connect request"));
    }
    const requestId = stringField(created.body, "requestId");
    const pollSecret = stringField(created.body, "pollSecret");
    const code = stringField(created.body, "code");
    const command = stringField(created.body, "command") || DESK_COMMAND;
    if (!requestId || !pollSecret || !code) throw new Error("Hermes returned an invalid connect request");
    onCode(code, command);
    for (;;) {
      if (signal?.aborted) throw new DOMException("Connect cancelled", "AbortError");
      await this.sleep(POLL_INTERVAL_MS);
      const polled = await this.post(base, `/${DESK_APP}/connect/status`, { requestId, pollSecret }, signal);
      if (polled.status !== 200) throw new Error(messageOf(polled.body, "Hermes did not answer the connect request"));
      const status = stringField(polled.body, "status");
      if (status === "pending") continue;
      if (status === "approved") {
        const grant = stringField(polled.body, "grant");
        if (!grant || grant.length < 32) throw new Error("Hermes returned an invalid approval");
        this.saveGrant(grant);
        return "approved";
      }
      return status === "denied" ? "denied" : "expired";
    }
  }

  private async post(
    base: string,
    path: string,
    body: unknown,
    signal?: AbortSignal,
  ): Promise<{ status: number; body: unknown }> {
    const response = await this.fetchImpl(`${base}${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      cache: "no-store",
      signal,
    });
    let parsed: unknown = null;
    try {
      parsed = await response.json();
    } catch {
      parsed = null;
    }
    return { status: response.status, body: parsed };
  }
}

function safeLocalStorage(): DeskStorage | null {
  try {
    return typeof localStorage === "undefined" ? null : localStorage;
  } catch {
    return null;
  }
}

function stringField(value: unknown, key: string): string {
  if (!value || typeof value !== "object") return "";
  const field = (value as Record<string, unknown>)[key];
  return typeof field === "string" ? field : "";
}

function messageOf(value: unknown, fallback: string): string {
  const message = stringField(value, "error");
  return message ? message.slice(0, 300) : fallback;
}
