// @vitest-environment node
/**
 * R4i: the task pane shows a bounded, content-free, accessible count of workbook actions that need
 * reconciliation (R3c D4 markers). Read-only and idle-only; no execution change.
 * - N1-N8: RED at the base (the module, the wiring, the HTML line and the CSS rules do not exist yet).
 * - G1-G3: GREEN pins that hold at the base and at HEAD.
 *
 * The notice module is new in R4i, so it is loaded through a variable specifier: at the base every
 * behaviour row fails with an AssertionError (missing export), never with a type or import error.
 * Store helpers are copied from browser-idempotency-store-r3c.test.ts as text; nothing is imported
 * from other test files.
 */
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  type IdempotencyStorage,
} from "../workbookPlanExecutor";

const MODULE_PATH = "../taskpane/reconciliationNotice";
let mod: Record<string, unknown> | undefined;
beforeAll(async () => {
  mod = await import(/* @vite-ignore */ MODULE_PATH).catch(() => undefined);
});

type View =
  | { readonly kind: "unavailable" }
  | { readonly kind: "none" }
  | { readonly kind: "show"; readonly text: string };
interface Target { textContent: string | null; hidden: boolean }
interface Source { reconciliationRequiredCount(): Promise<number | undefined> }
interface Notice { refresh(): Promise<void> }
interface NoticeApi {
  readonly RECONCILIATION_NOTICE_MAX_DISPLAY: number;
  reconciliationNoticeView(count: unknown): View;
  createReconciliationNotice(deps: {
    readonly source: Source;
    readonly target: () => Target | null;
    readonly isIdle: () => boolean;
  }): Notice;
}

function api(): NoticeApi {
  expect(typeof mod?.reconciliationNoticeView, "reconciliationNoticeView export").toBe("function");
  expect(typeof mod?.createReconciliationNotice, "createReconciliationNotice export").toBe("function");
  return mod as unknown as NoticeApi;
}

const TEXT_SHAPE = /^(?:[1-9]\d{0,2}|999\+) workbook actions? needs? reconciliation$/;
const ONE = "1 workbook action needs reconciliation";
const plural = (n: number | string): string => `${n} workbook actions need reconciliation`;

const PANE_DIR = resolve(__dirname, "../taskpane");
const MODULE_FILE = join(PANE_DIR, "reconciliationNotice.ts");
const paneTs = (): string => readFileSync(join(PANE_DIR, "taskpane.ts"), "utf8");
const paneHtml = (): string => readFileSync(join(PANE_DIR, "taskpane.html"), "utf8");
const paneCss = (): string => readFileSync(join(PANE_DIR, "taskpane.css"), "utf8");

/** Removes block comments and whole-line // comments (taskpane.ts has no comment markers inside strings). */
function stripComments(source: string): string {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .split("\n")
    .filter((line) => !/^\s*\/\//.test(line))
    .join("\n");
}

const occurrences = (text: string, needle: string): number => text.split(needle).length - 1;

/** Slice from the anchor to the first `terminator` after it (inclusive). */
function slice(text: string, anchor: string, terminator: string): string {
  const start = text.indexOf(anchor);
  expect(start, `anchor ${anchor}`).toBeGreaterThan(-1);
  const end = text.indexOf(terminator, start);
  expect(end, `terminator after ${anchor}`).toBeGreaterThan(start);
  return text.slice(start, end + terminator.length);
}

function deferred<T>(): { readonly promise: Promise<T>; resolve(value: T): void } {
  let resolveFn!: (value: T) => void;
  const promise = new Promise<T>((r) => { resolveFn = r; });
  return { promise, resolve: (value) => resolveFn(value) };
}

// ---------------------------------------------------------------------------
// Store helpers (copied from browser-idempotency-store-r3c.test.ts, not imported)
// ---------------------------------------------------------------------------
const H = 10 * 60 * 1000;
const STORAGE_KEY = "hermes.excel.workbookPlan.idempotency.v1";
const SESSION = "5e550000-0000-4000-8000-000000000001";
const DOCUMENT = "d0c00000-0000-4000-8000-000000000002";
const uuid = (n: number): string => `00000000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const turnUuid = (n: number): string => `7a7a0000-0000-4000-8000-${n.toString(16).padStart(12, "0")}`;
const tokenFor = (turn: number): string => `${SESSION}.${DOCUMENT}.${turnUuid(turn)}`;
const identity = (n: number): { readonly boundary: string; readonly tx: string; readonly digest: string } => ({
  boundary: tokenFor(n),
  tx: uuid(0x100000 + n),
  digest: `digest-${n}`,
});

class FakeLocalStorage implements IdempotencyStorage {
  private readonly values = new Map<string, string>();
  readonly writes: { readonly key: string; readonly value: string }[] = [];
  removals = 0;
  throwOnGet = false;

  getItem(key: string): string | null {
    if (this.throwOnGet) throw new Error("storage read failed");
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.writes.push({ key, value });
    this.values.set(key, value);
  }

  removeItem(key: string): void {
    this.removals += 1;
    this.values.delete(key);
  }

  raw(): string | null {
    return this.values.get(STORAGE_KEY) ?? null;
  }

  records(): Record<string, unknown>[] {
    const raw = this.raw();
    return raw === null ? [] : JSON.parse(raw) as Record<string, unknown>[];
  }
}

const claimId = (store: BrowserIdempotencyStore, n: number) => {
  const id = identity(n);
  return store.claim(id.boundary, id.tx, id.digest);
};

/** Orphans `orphans` (claimed at t=0 by store A), then store B claims `trigger` at t=H+1, which marks them. */
async function seededMarkers(orphans: readonly number[], trigger: number): Promise<{
  readonly storage: FakeLocalStorage;
  readonly store: BrowserIdempotencyStore;
}> {
  let now = 0;
  const storage = new FakeLocalStorage();
  const a = new BrowserIdempotencyStore(storage, { clock: () => now });
  for (const n of orphans) expect((await claimId(a, n)).kind).toBe("claimed");
  now = H + 1;
  const b = new BrowserIdempotencyStore(storage, { clock: () => now });
  expect((await claimId(b, trigger)).kind).toBe("claimed");
  return { storage, store: b };
}

// ---------------------------------------------------------------------------
// RED rows
// ---------------------------------------------------------------------------
describe("R4i N1 content-free view table", () => {
  it("N1 maps counts to unavailable / none / singular / plural / 999+ text", () => {
    const { reconciliationNoticeView, RECONCILIATION_NOTICE_MAX_DISPLAY } = api();
    expect(RECONCILIATION_NOTICE_MAX_DISPLAY).toBe(999);
    for (const bad of [undefined, null, Number.NaN, Number.POSITIVE_INFINITY, -1, 1.5, "3", 2 ** 53, {}, true]) {
      expect(reconciliationNoticeView(bad), `view(${String(bad)})`).toEqual({ kind: "unavailable" });
    }
    expect(reconciliationNoticeView(0)).toEqual({ kind: "none" });
    expect(reconciliationNoticeView(1)).toEqual({ kind: "show", text: ONE });
    for (const n of [2, 127, 128, 999]) expect(reconciliationNoticeView(n)).toEqual({ kind: "show", text: plural(n) });
    for (const n of [1000, 5000]) expect(reconciliationNoticeView(n)).toEqual({ kind: "show", text: plural("999+") });
    for (const n of [1, 2, 9, 10, 99, 100, 127, 128, 999, 1000, 5000, 2 ** 31]) {
      const view = reconciliationNoticeView(n);
      expect(view.kind).toBe("show");
      const text = view.kind === "show" ? view.text : "";
      expect(text).toMatch(TEXT_SHAPE);
      expect(text.length).toBeLessThanOrEqual(45);
      expect(/ action needs /.test(text)).toBe(n === 1);
    }
  });
});

describe("R4i N2 refresh semantics (fake source and target)", () => {
  const setup = (opts: { idle?: () => boolean; target?: () => Target | null } = {}) => {
    const target: Target = { textContent: "", hidden: true };
    let next: () => Promise<number | undefined> = async () => undefined;
    let calls = 0;
    const source: Source = {
      reconciliationRequiredCount: () => {
        calls += 1;
        return next();
      },
    };
    const notice = api().createReconciliationNotice({
      source,
      target: opts.target ?? (() => target),
      isIdle: opts.idle ?? (() => true),
    });
    return {
      target,
      notice,
      calls: () => calls,
      answer: (fn: () => Promise<number | undefined>) => { next = fn; },
    };
  };

  it("N2a-c shows a positive count, hides on 0, and leaves the display unchanged on undefined", async () => {
    const s = setup();
    s.answer(async () => 3);
    await s.notice.refresh();
    expect(s.target).toEqual({ textContent: plural(3), hidden: false });
    s.answer(async () => 0);
    await s.notice.refresh();
    expect(s.target).toEqual({ textContent: "", hidden: true });
    s.answer(async () => 5);
    await s.notice.refresh();
    expect(s.target).toEqual({ textContent: plural(5), hidden: false });
    s.answer(async () => undefined);
    await s.notice.refresh();
    expect(s.target).toEqual({ textContent: plural(5), hidden: false });
    for (const bad of [Number.NaN, -1, 1.5] as number[]) {
      s.answer(async () => bad);
      await s.notice.refresh();
      expect(s.target).toEqual({ textContent: plural(5), hidden: false });
    }
    expect(s.calls()).toBe(7);
  });

  it("N2d a synchronous throw or a rejection leaves the display unchanged and refresh still resolves", async () => {
    const s = setup();
    s.answer(async () => 1);
    await s.notice.refresh();
    expect(s.target).toEqual({ textContent: ONE, hidden: false });
    s.answer(() => { throw new Error("sync boom"); });
    let first: Promise<void> | undefined;
    expect(() => { first = s.notice.refresh(); }).not.toThrow();
    await expect(first).resolves.toBeUndefined();
    expect(s.target).toEqual({ textContent: ONE, hidden: false });
    s.answer(() => Promise.reject(new Error("async boom")));
    await expect(s.notice.refresh()).resolves.toBeUndefined();
    expect(s.target).toEqual({ textContent: ONE, hidden: false });
    expect(s.calls()).toBe(3);
  });

  it("N2e when the pane is not idle the source is not called and the target is untouched", async () => {
    let idle = false;
    let targetReads = 0;
    const target: Target = { textContent: "keep", hidden: false };
    const s = setup({ idle: () => idle, target: () => { targetReads += 1; return target; } });
    s.answer(async () => 0);
    await expect(s.notice.refresh()).resolves.toBeUndefined();
    expect(s.calls()).toBe(0);
    expect(targetReads).toBe(0);
    expect(target).toEqual({ textContent: "keep", hidden: false });
    idle = true;
    await s.notice.refresh();
    expect(s.calls()).toBe(1);
    expect(target).toEqual({ textContent: "", hidden: true });
  });

  it("N2e' an isIdle that throws is treated as not idle and refresh resolves", async () => {
    const s = setup({ idle: () => { throw new Error("idle boom"); } });
    s.answer(async () => 2);
    let p: Promise<void> | undefined;
    expect(() => { p = s.notice.refresh(); }).not.toThrow();
    await expect(p).resolves.toBeUndefined();
    expect(s.calls()).toBe(0);
    expect(s.target).toEqual({ textContent: "", hidden: true });
  });

  it("N2f single-flight: calls while a read is pending share it; the next call after it settles reads again", async () => {
    const s = setup();
    const gate = deferred<number | undefined>();
    s.answer(() => gate.promise);
    const p1 = s.notice.refresh();
    const p2 = s.notice.refresh();
    const p3 = s.notice.refresh();
    expect(s.calls()).toBe(1);
    expect(p2).toBe(p1);
    expect(p3).toBe(p1);
    gate.resolve(4);
    await expect(Promise.all([p1, p2, p3])).resolves.toEqual([undefined, undefined, undefined]);
    expect(s.target).toEqual({ textContent: plural(4), hidden: false });
    s.answer(async () => 0);
    await s.notice.refresh();
    // R4j amendment (R4i lane-B B-A1): the two joined calls add exactly one coalesced follow-up read (2 -> 3).
    expect(s.calls()).toBe(3);
    expect(s.target).toEqual({ textContent: "", hidden: true });
  });

  it("N2g a missing target element is tolerated", async () => {
    const s = setup({ target: () => null });
    s.answer(async () => 7);
    await expect(s.notice.refresh()).resolves.toBeUndefined();
    expect(s.calls()).toBe(1);
  });
});

describe("R4i N3/N4 real BrowserIdempotencyStore: no writes, content-free text", () => {
  it("N3 one orphan marker shows the singular line, the refresh writes nothing, and a read failure keeps the display", async () => {
    const { storage, store } = await seededMarkers([1], 2);
    const target: Target = { textContent: "", hidden: true };
    const notice = api().createReconciliationNotice({ source: store, target: () => target, isIdle: () => true });
    const writesBefore = storage.writes.length;
    const removalsBefore = storage.removals;
    const rawBefore = storage.raw();
    expect(rawBefore).not.toBeNull();
    await notice.refresh();
    expect(target).toEqual({ textContent: ONE, hidden: false });
    expect(storage.writes.length).toBe(writesBefore);
    expect(storage.removals).toBe(removalsBefore);
    expect(storage.raw()).toBe(rawBefore);

    storage.throwOnGet = true;
    await expect(notice.refresh()).resolves.toBeUndefined();
    expect(target).toEqual({ textContent: ONE, hidden: false });
    expect(storage.writes.length).toBe(writesBefore);
  });

  it("N4 the text carries no record content; two markers give the plural line", async () => {
    const { storage, store } = await seededMarkers([1], 2);
    const target: Target = { textContent: "", hidden: true };
    await api().createReconciliationNotice({ source: store, target: () => target, isIdle: () => true }).refresh();
    expect(target.textContent).toBe(ONE);
    const forbidden = new Set<string>(["restart_rejected", SESSION, DOCUMENT]);
    for (const n of [1, 2]) {
      const id = identity(n);
      forbidden.add(id.boundary).add(id.tx).add(id.digest).add(turnUuid(n));
    }
    for (const record of storage.records()) {
      for (const field of ["boundary", "transactionId", "digest", "owner", "fence"]) {
        const value = record[field];
        if (typeof value === "string" && value.length > 0) forbidden.add(value);
      }
    }
    expect(forbidden.size).toBeGreaterThan(8);
    for (const value of forbidden) expect(target.textContent ?? "").not.toContain(value);

    const two = await seededMarkers([1, 2], 3);
    expect(await two.store.reconciliationRequiredCount()).toBe(2);
    const target2: Target = { textContent: "", hidden: true };
    await api().createReconciliationNotice({ source: two.store, target: () => target2, isIdle: () => true }).refresh();
    expect(target2).toEqual({ textContent: plural(2), hidden: false });
  });
});

describe("R4i N5/N6 accessible status line (static HTML/CSS)", () => {
  it("N5 one hidden role=status line between the bridge strip and #messages, with no data or static text", () => {
    const html = paneHtml();
    expect(occurrences(html, 'id="reconciliation-status"')).toBe(1);
    const tag = html.match(/<p\b([^>]*\bid="reconciliation-status"[^>]*)>([\s\S]*?)<\/p>/);
    expect(tag, "reconciliation-status <p> element").not.toBeNull();
    const attrs = tag?.[1] ?? "";
    expect(attrs).toMatch(/\bclass="reconciliation-status"/);
    expect(attrs).toMatch(/\brole="status"/);
    expect(attrs).toMatch(/\baria-live="polite"/);
    expect(attrs).toMatch(/\baria-atomic="true"/);
    expect(attrs).toMatch(/(?:^|\s)hidden(?:\s|$)/);
    expect(attrs).not.toMatch(/data-|title=|tabindex=|onclick/i);
    expect(tag?.[2]).toBe("");
    const strip = html.indexOf('<section class="bridge-strip" aria-live="polite">');
    const stripEnd = html.indexOf("</section>", strip);
    const at = html.indexOf('id="reconciliation-status"');
    const messages = html.indexOf('id="messages"');
    expect(strip).toBeGreaterThan(-1);
    expect(stripEnd).toBeGreaterThan(strip);
    expect(at).toBeGreaterThan(stripEnd);
    expect(messages).toBeGreaterThan(at);
  });

  it("N6 the line has its own rule with the checkpoint palette and a [hidden] display:none override", () => {
    const css = paneCss();
    const rule = (selector: string): string => {
      const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const matches = [...css.matchAll(new RegExp(`(?:^|[\\n}])\\s*${escaped}\\s*\\{([^}]*)\\}`, "g"))];
      return matches.at(-1)?.[1] ?? "";
    };
    const base = rule(".reconciliation-status");
    expect(base).toMatch(/(?:^|[\s;])color:\s*var\(--caution\)\s*;/);
    expect(base).toMatch(/background:\s*var\(--raised\)\s*;/);
    expect(rule(".reconciliation-status[hidden]")).toMatch(/display:\s*none\s*;/);
  });
});

describe("R4i N7 taskpane.ts wiring (source scan, comments stripped)", () => {
  it("N7 imports once, builds the notice after the route, and refreshes exactly at onReady and finishActiveTurn", () => {
    const raw = paneTs();
    const src = stripComments(raw);
    expect(occurrences(src, 'import { createReconciliationNotice } from "./reconciliationNotice";')).toBe(1);
    expect(occurrences(src, "createReconciliationNotice(")).toBe(1);
    expect(occurrences(src, "void reconciliationNotice.refresh();")).toBe(2);
    expect(occurrences(src, "reconciliationNotice.refresh(")).toBe(2);
    expect(occurrences(src, "await reconciliationNotice")).toBe(0);

    const onReady = slice(src, "  Office.onReady(() => {", "\n  });");
    expect(onReady).toContain(
      "    updateBridgeState({ connected: false, sessionId: null, documentId: null });\n    void reconciliationNotice.refresh();\n",
    );
    const finish = slice(src, "function finishActiveTurn(", "\n}\n");
    expect(finish.endsWith("  currentAssistantMessage = null;\n  setRunning(false);\n  void reconciliationNotice.refresh();\n}\n")).toBe(true);
    expect(occurrences(onReady, "reconciliationNotice.refresh(")).toBe(1);
    expect(occurrences(finish, "reconciliationNotice.refresh(")).toBe(1);
    const stop = slice(src, "function stopRunningAgent(", "\n}\n");
    expect(stop).not.toContain("reconciliationNotice");
    const send = slice(src, "async function sendPrompt(", "\n}\n");
    expect(send).not.toContain("reconciliationNotice");

    const lines = src.split("\n");
    const countLines = lines.filter((line) => line.includes("reconciliationRequiredCount"));
    expect(countLines).toEqual([
      "  source: { reconciliationRequiredCount: () => new BrowserIdempotencyStore().reconciliationRequiredCount() },",
    ]);
    expect(lines.filter((line) => line.includes("isIdle:"))).toEqual([
      "  isIdle: () => abortController === null && activeTurnOperationId === null,",
    ]);

    const routeEnd = src.indexOf("  executePrimitive: serverPrimitiveGate,\n});\n");
    const decl = src.indexOf("const reconciliationNotice = createReconciliationNotice({");
    const declRaw = raw.indexOf("const reconciliationNotice = createReconciliationNotice({");
    expect(routeEnd).toBeGreaterThan(-1);
    expect(decl).toBeGreaterThan(routeEnd);
    expect(declRaw).toBeGreaterThan(raw.indexOf("  executePrimitive: serverPrimitiveGate,\n});\n"));
    expect(raw.indexOf("// Checkpoint UI state")).toBeGreaterThan(declRaw);
    expect(slice(src, "const reconciliationNotice = createReconciliationNotice({", "\n});").split("\n")).toEqual([
      "const reconciliationNotice = createReconciliationNotice({",
      "  source: { reconciliationRequiredCount: () => new BrowserIdempotencyStore().reconciliationRequiredCount() },",
      '  target: () => document.getElementById("reconciliation-status"),',
      "  isIdle: () => abortController === null && activeTurnOperationId === null,",
      "});",
    ]);
    expect(raw).toContain(
      "  executePrimitive: serverPrimitiveGate,\n});\n// R4i (R3c D4): content-free marker count, read only while this pane is idle. Never awaited by a turn.\nconst reconciliationNotice = createReconciliationNotice({\n",
    );

    // The R4d server primitive gate block keeps its exact 7-line shape (server-replay-deadline-r4e W1).
    const gate = slice(raw, "const serverPrimitiveGate = createServerPrimitiveGate({", "});");
    expect(gate.split("\n")).toHaveLength(7);
    expect(gate).not.toContain("reconciliation");
  });
});

describe("R4i N8 notice module hygiene (source scan)", () => {
  it("N8 has no value import, no store/claim/write/HTML/timer token, no capacity token and no authority scan hit", () => {
    expect(existsSync(MODULE_FILE), "src/taskpane/reconciliationNotice.ts exists").toBe(true);
    const text = existsSync(MODULE_FILE) ? readFileSync(MODULE_FILE, "utf8") : "";
    const imports = text.match(/^\s*import\b.*$/gm) ?? [];
    for (const line of imports) expect(line).toMatch(/^\s*import\s+type\s/);
    expect(text).not.toMatch(/^\s*export\s+[^;]*\bfrom\s/m);
    expect(text).not.toMatch(/\brequire\(|\bimport\(/);
    for (const token of [
      "claim(", "finish(", "setItem", "removeItem", "localStorage", "innerHTML", "insertAdjacentHTML", "outerHTML",
      ".execute(", "setTimeout", "setInterval", "plan_fence_capacity", "BrowserIdempotencyStore",
      "installFullFunctionalityAcceptanceHarness", "__hermesExcelFullFunctionalityAcceptance", "Hermes_Acceptance_",
    ]) {
      expect(text, token).not.toContain(token);
    }
    for (const pattern of [
      /function\s+claimRawPrimitiveExecutor\s*\(/, /(?<!function )\bclaimRawPrimitiveExecutor\s*\(/, /\bclaimRawPrimitiveExecutor\b/,
      /\.execute\(/, /\bgetExcelToolDefinition\(/, /\bexecuteLegacyServerPrimitiveR4dPending\b/, /\bfunction\s+createServerPrimitiveGate\(/,
      /(?<!function\s)\bcreateServerPrimitiveGate\(/, /\bserverPrimitiveAuthority\(/, /\.serverPrimitiveAuthority\(/,
      /\bcreatePrimitiveMutationAuthority\(/, /\bnew\s+PrimitiveMutationAuthority\b/, /PRIMITIVE_EFFECTS\[/, /\bPRIMITIVE_EFFECTS\b/,
    ]) {
      expect(text, String(pattern)).not.toMatch(pattern);
    }
    expect(text).toContain("textContent");
  });
});

// ---------------------------------------------------------------------------
// GREEN pins (hold at the base and at HEAD)
// ---------------------------------------------------------------------------
describe("R4i GREEN pins", () => {
  it("G1 the store contract R4i relies on: reconciliationRequiredCount exists and is 0 over empty storage", async () => {
    expect(typeof BrowserIdempotencyStore.prototype.reconciliationRequiredCount).toBe("function");
    await expect(new BrowserIdempotencyStore(new MemoryIdempotencyStorage()).reconciliationRequiredCount()).resolves.toBe(0);
  });

  it("G2 taskpane.ts keeps exactly one finishActiveTurn and one Office.onReady", () => {
    const src = paneTs();
    expect(occurrences(src, "function finishActiveTurn(")).toBe(1);
    expect(occurrences(src, "Office.onReady(")).toBe(1);
  });

  it("G3 no plan_fence_capacity token anywhere in src/taskpane (capacity signal deferred)", () => {
    const hits: string[] = [];
    const walk = (dir: string): void => {
      for (const entry of readdirSync(dir).sort()) {
        const full = join(dir, entry);
        if (statSync(full).isDirectory()) walk(full);
        else if (readFileSync(full, "utf8").includes("plan_fence_capacity")) hits.push(entry);
      }
    };
    walk(PANE_DIR);
    expect(hits).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// R4j fold-in (R4i lane-B B-A1): a refresh that joins a pending read schedules exactly one
// coalesced follow-up read after it settles, so the line ends at the latest count. Still
// single-flight (never two reads at once) and idle-only. NJ1-NJ3 and NJ6 are RED at f3e3c0a;
// NJ4 and NJ5 are pins (identical at f3e3c0a and HEAD).
// ---------------------------------------------------------------------------
describe("R4j B-A1 coalesced follow-up read", () => {
  const counted = (opts: { idle?: () => boolean } = {}) => {
    const target: Target = { textContent: "", hidden: true };
    const answers: (() => Promise<number | undefined>)[] = [];
    let calls = 0;
    let inFlight = 0;
    let maxInFlight = 0;
    const source: Source = {
      reconciliationRequiredCount: () => {
        calls += 1;
        inFlight += 1;
        maxInFlight = Math.max(maxInFlight, inFlight);
        const next = answers.shift() ?? (async () => undefined);
        return next().finally(() => { inFlight -= 1; });
      },
    };
    const notice = api().createReconciliationNotice({ source, target: () => target, isIdle: opts.idle ?? (() => true) });
    return { target, notice, answers, calls: () => calls, maxInFlight: () => maxInFlight };
  };

  it("NJ1 a refresh joining a pending read gets one follow-up read; the joined promise settles after it and the line shows the latest count", async () => {
    const s = counted();
    const gate = deferred<number | undefined>();
    s.answers.push(() => gate.promise, async () => 9);
    const p1 = s.notice.refresh();
    const p2 = s.notice.refresh();
    expect(p2).toBe(p1);
    expect(s.calls()).toBe(1);
    gate.resolve(2);
    await p2;
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 2, target: { textContent: plural(9), hidden: false } });
  });

  it("NJ2 many joins coalesce into exactly one follow-up read (never more than one queued)", async () => {
    const s = counted();
    const gate = deferred<number | undefined>();
    s.answers.push(() => gate.promise, async () => 4, async () => 5, async () => 6);
    const first = s.notice.refresh();
    const joined = [1, 2, 3, 4, 5].map(() => s.notice.refresh());
    for (const p of joined) expect(p).toBe(first);
    gate.resolve(1);
    await Promise.all([first, ...joined]);
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 2, target: { textContent: plural(4), hidden: false } });
    await s.notice.refresh();
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 3, target: { textContent: plural(5), hidden: false } });
  });

  it("NJ3 single-flight holds through the follow-up: it starts only after the first read settles; a join during it queues one more", async () => {
    const s = counted();
    const g1 = deferred<number | undefined>();
    const g2 = deferred<number | undefined>();
    s.answers.push(() => g1.promise, () => g2.promise, async () => 0);
    const p = s.notice.refresh();
    expect(s.notice.refresh()).toBe(p);
    expect(s.calls()).toBe(1);
    g1.resolve(3);
    for (let i = 0; i < 20 && s.calls() < 2; i += 1) await Promise.resolve();
    expect(s.calls()).toBe(2);
    expect(s.target).toEqual({ textContent: plural(3), hidden: false });
    expect(s.notice.refresh()).toBe(p);
    expect(s.notice.refresh()).toBe(p);
    expect(s.calls()).toBe(2);
    g2.resolve(7);
    await p;
    expect({ calls: s.calls(), maxInFlight: s.maxInFlight(), target: s.target })
      .toEqual({ calls: 3, maxInFlight: 1, target: { textContent: "", hidden: true } });
  });

  it("NJ4 [pin] the follow-up is idle-only: if the pane is busy when the first read settles, no follow-up read happens", async () => {
    let idle = true;
    const s = counted({ idle: () => idle });
    const gate = deferred<number | undefined>();
    s.answers.push(() => gate.promise, async () => 9);
    const p = s.notice.refresh();
    void s.notice.refresh();
    idle = false;
    gate.resolve(2);
    await p;
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 1, target: { textContent: plural(2), hidden: false } });
    idle = true;
    await s.notice.refresh();
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 2, target: { textContent: plural(9), hidden: false } });
  });

  it("NJ5 [pin] a lone refresh (no join) reads exactly once", async () => {
    const s = counted();
    s.answers.push(async () => 1, async () => 5);
    await s.notice.refresh();
    await Promise.resolve();
    expect({ calls: s.calls(), maxInFlight: s.maxInFlight(), target: s.target }).toEqual({ calls: 1, maxInFlight: 1, target: { textContent: ONE, hidden: false } });
  });

  it("NJ6 a failed follow-up read keeps the display, refresh still resolves, and the next call reads again", async () => {
    const s = counted();
    const gate = deferred<number | undefined>();
    s.answers.push(() => gate.promise, () => Promise.reject(new Error("follow-up boom")), async () => 0);
    const p = s.notice.refresh();
    void s.notice.refresh();
    gate.resolve(6);
    await expect(p).resolves.toBeUndefined();
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 2, target: { textContent: plural(6), hidden: false } });
    await s.notice.refresh();
    expect({ calls: s.calls(), target: s.target }).toEqual({ calls: 3, target: { textContent: "", hidden: true } });
  });
});
