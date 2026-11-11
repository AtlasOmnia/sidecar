#!/usr/bin/env node
/** Offline synthetic evaluator. Only the Office host is doubled; kernel modules are real. */
import { createHash, randomBytes } from "node:crypto";
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { existsSync, lstatSync, readFileSync, readdirSync, mkdirSync, writeFileSync, realpathSync, mkdtempSync, symlinkSync } from "node:fs";
import { dirname, basename, join, resolve, relative, isAbsolute, sep } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
export const EVALUATOR_VERSION = "1.2.0";
const hash = data => createHash("sha256").update(data).digest("hex");
const clone = value => structuredClone(value);
const canonical = value => {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === "object") return Object.fromEntries(Object.keys(value).sort().map(k => [k, canonical(value[k])]));
  return value;
};
const ids = { sessionId: "11111111-1111-4111-8111-111111111111", documentId: "22222222-2222-4222-8222-222222222222", turnId: "33333333-3333-4333-8333-333333333333" };
const originalHash = "fa2bae32f237be940024762f78ec9d9f1c9b9e2ba97df0ae292e444009431cbf";
const sensitiveKeyNames = "token|access[_-]?token|refresh[_-]?token|api[_-]?key|uuid|bridgeId|sessionId|documentId|turnId|secret|password|credential|authorization";
const sensitivePatterns = [/\/Users\//, /\/home\//, /\/C:\\/, /\\Users\\/, /sk-[A-Za-z0-9]{20,}/, /api[_-]?key/i, /token[=:\s]/i, new RegExp(`"(?:${sensitiveKeyNames})"\\s*:`, "i"), /uuid[:=]/i, /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i, /bridgeId/i, /\bSyntheticOrg\b/i, /\bPrivateOrgSentinel\b/i, /\bpayroll\b/i, /\bclaim[s]?\b/i, /\bhealth\b/i, /\bpersonal\b/i, /\bsocial.security\b/i, /\bmedicare\b/i, /\bdiagnosis\b/i, /\bclaimNumber\b/i];
const sensitiveKeyPatterns = [new RegExp(`^(?:${sensitiveKeyNames})$`, "i")];
function decodeEscapes(value) {
  return value.replace(/\\u([0-9a-fA-F]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/\\\//g, "/").replace(/\\"/g, '"').replace(/\\\\/g, "\\");
}
// Extract balanced top-level JSON object/array candidates from a transport
// string (handles multiple envelopes, nested braces, and string literals).
function extractJsonCandidates(value) {
  const out = [];
  let i = 0;
  while (i < value.length) {
    const open = value[i];
    if (open !== "{" && open !== "[") { i++; continue; }
    const close = open === "{" ? "}" : "]";
    let depth = 0, inStr = false, esc = false, j = i;
    for (; j < value.length; j++) {
      const ch = value[j];
      if (inStr) { if (esc) esc = false; else if (ch === "\\") esc = true; else if (ch === '"') inStr = false; continue; }
      if (ch === '"') { inStr = true; continue; }
      if (ch === open) depth++;
      else if (ch === close) { depth--; if (depth === 0) { out.push(value.slice(i, j + 1)); i = j + 1; break; } }
    }
    if (j >= value.length) break; // unbalanced: stop; the remainder is nested inside an unclosed delimiter
  }
  return out;
}
function assertPrivacyClean(value, seen = new Set(), depth = 0) {
  if (value === null || value === undefined) return;
  if (typeof value === "string") {
    if (sensitivePatterns.some(p => p.test(value))) throw coded("corpus_sensitive_pattern");
    // Fail closed on inputs we cannot fully inspect: oversized payloads and
    // excessive nesting must be rejected, never treated as clean.
    if (depth >= 32 || value.length > 262144) throw coded("corpus_inspection_limit_exceeded");
    // Recursively decode escape layers: each pass can reveal a deeper JSON
    // layer (nested escaped quotes) that a single pass would miss.
    const unescaped = decodeEscapes(value);
    if (unescaped !== value) {
      try { assertPrivacyClean(unescaped, seen, depth + 1); }
      catch (error) { if (error?.code === "corpus_sensitive_pattern" || error?.code === "corpus_inspection_limit_exceeded") throw error; }
    }
    const trimmed = value.trim();
    if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
      try { assertPrivacyClean(JSON.parse(trimmed), seen, depth + 1); }
      catch (error) { if (error?.code === "corpus_sensitive_pattern" || error?.code === "corpus_inspection_limit_exceeded") throw error; }
    }
    // Transport envelopes (XML/OpenAI) can carry one or more JSON payloads;
    // inspect each individually rather than first-to-last delimiter slicing.
    for (const candidate of extractJsonCandidates(value)) {
      try { assertPrivacyClean(JSON.parse(candidate), seen, depth + 1); }
      catch (error) { if (error?.code === "corpus_sensitive_pattern" || error?.code === "corpus_inspection_limit_exceeded") throw error; }
    }
    return;
  }
  if (typeof value !== "object") return;
  if (seen.has(value)) return; seen.add(value);
  for (const [key, child] of Object.entries(value)) {
    if (typeof key === "string") {
      if (sensitiveKeyPatterns.some(p => p.test(key))) throw coded("corpus_sensitive_pattern");
      // Sensitive content in an ordinary property name is still a disclosure.
      if (sensitivePatterns.some(p => p.test(key))) throw coded("corpus_sensitive_pattern");
      // Escaped property names can hide sensitive patterns; decode and re-check.
      const unescapedKey = decodeEscapes(key);
      if (unescapedKey !== key && (sensitiveKeyPatterns.some(p => p.test(unescapedKey)) || sensitivePatterns.some(p => p.test(unescapedKey)))) throw coded("corpus_sensitive_pattern");
    }
    assertPrivacyClean(child, seen, depth);
  }
}
function coded(code, exitCode = 3) { return Object.assign(new Error(code), { code, exitCode }); }
function safeMetadata(value, fallback) {
  return typeof value === "string" && /^[a-zA-Z0-9_.:-]{1,100}$/.test(value) ? value : fallback;
}
// Candidate identity is written to every evidence record; constrain it to an
// exact SHA so arbitrary (possibly secret) metadata cannot leak into records.
function validateCandidateSha(value) {
  if (value === undefined || value === null || value === "") return "unknown";
  if (value === "unknown") return "unknown";
  if (typeof value !== "string" || !/^[0-9a-f]{40}$/.test(value)) throw coded("invalid_candidate_sha", 2);
  return value;
}
function realDestination(path) {
  let ancestor = path; const suffix = [];
  // basename is separator-aware (both `/` and `\` on Windows), so forward-slash
  // aliases cannot smuggle a resolved ancestor past the containment check.
  while (!existsSync(ancestor)) { suffix.unshift(basename(ancestor)); const parent = dirname(ancestor); if (parent === ancestor) throw coded("invalid_out", 2); ancestor = parent; }
  return resolve(realpathSync(ancestor), ...suffix);
}
export function validateOutputDirectory(out) {
  if (typeof out !== "string" || !isAbsolute(out)) throw coded("absolute_external_out_required", 2);
  const target = realDestination(out); const rel = relative(realpathSync(repoRoot), target);
  if (rel === "" || (!rel.startsWith(`..${sep}`) && rel !== ".." && !isAbsolute(rel))) throw coded("repo_output_forbidden", 2);
  const recordStat = lstatSync(join(target, "records.jsonl"), { throwIfNoEntry: false });
  if (recordStat && (!recordStat.isFile() || recordStat.nlink !== 1)) throw coded("output_file_alias_forbidden", 2);
  return target;
}
export function validateScratchDirectory(scratch) {
  if (typeof scratch !== "string" || !isAbsolute(scratch)) throw coded("scratch_directory_required", 2);
  const target = realDestination(scratch); const rel = relative(realpathSync(repoRoot), target);
  if (rel === "" || (!rel.startsWith(`..${sep}`) && rel !== ".." && !isAbsolute(rel))) throw coded("repo_scratch_forbidden", 2);
  if (!lstatSync(target, { throwIfNoEntry: false })?.isDirectory()) throw coded("scratch_directory_required", 2);
  return target;
}
function validateManifestMetadata(manifest) {
  if (!manifest || typeof manifest !== "object" || Array.isArray(manifest)
    || (manifest.corpusId !== undefined && !["workbook-plan-conformance", "workbook-plan-shadow"].includes(manifest.corpusId))
    || typeof manifest.corpusVersion !== "string" || !/^\d{1,6}\.\d{1,6}\.\d{1,6}$/.test(manifest.corpusVersion)
    || typeof manifest.corpusSha256 !== "string" || !/^[0-9a-f]{64}$/.test(manifest.corpusSha256)
    || !Number.isInteger(manifest.fixtureCount) || manifest.fixtureCount < 1 || manifest.fixtureCount > 1000) throw coded("invalid_corpus_metadata");
}
export function loadCorpus(directory, options = {}) {
  const manifest = JSON.parse(readFileSync(join(directory, "manifest.json"), "utf8"));
  validateManifestMetadata(manifest);
  const role = options.role ?? (manifest.corpusId === undefined ? "original" : "synthetic");
  if (!["original", "synthetic"].includes(role)) throw coded("invalid_corpus_role");
  if (role === "original" && (manifest.corpusId !== undefined || manifest.corpusSha256 !== originalHash || manifest.fixtureCount !== 12)) throw coded("frozen_original_required");
  if (role === "synthetic" && manifest.corpusId === undefined) throw coded("synthetic_corpus_id_required");
  const names = readdirSync(directory).filter(n => n.endsWith(".json") && n !== "manifest.json").sort();
  if (!manifest.files || names.length !== manifest.fixtureCount || JSON.stringify(names) !== JSON.stringify(Object.keys(manifest.files).sort())) throw coded("corpus_inventory_mismatch");
  const fixtures = []; const digests = [];
  for (const name of names) {
    const bytes = readFileSync(join(directory, name)); const digest = hash(bytes);
    if (digest !== manifest.files[name]) throw coded("corpus_hash_mismatch");
    const fixture = JSON.parse(bytes.toString("utf8"));
    if (role === "synthetic") assertPrivacyClean(fixture);
    if (!/^[a-zA-Z0-9_-]{1,80}$/.test(fixture.caseId)) throw coded("invalid_case_id");
    fixtures.push(fixture); digests.push(`${name}:${digest}`);
  }
  const aggregate = hash(digests.join(""));
  if (aggregate !== manifest.corpusSha256 || (role === "original" && (aggregate !== originalHash || names.length !== 12))) throw coded("corpus_aggregate_mismatch");
  if (new Set(fixtures.map(f => f.caseId)).size !== fixtures.length) throw coded("duplicate_case_id");
  return { role, manifest, fixtures };
}
function bounds(address) {
  const local = address.split("!").at(-1).replace(/\$/g, "").toUpperCase();
  const match = /^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/.exec(local);
  if (!match) throw coded("unsupported_range");
  const col = letters => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const start = { row: Number(match[2]), col: col(match[1]) }; const end = { row: Number(match[4] ?? match[2]), col: col(match[3] ?? match[1]) };
  if (start.row < 1 || start.col < 1 || end.row > 1048576 || end.col > 16384) throw coded("double_range_limit");
  if (end.row < start.row || end.col < start.col || (end.row - start.row + 1) * (end.col - start.col + 1) > 100000) throw coded("double_range_limit");
  return { local, start, rows: end.row - start.row + 1, cols: end.col - start.col + 1 };
}
function column(n) { let s = ""; while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); } return s; }
/** Every proxy is batch-local: loads snapshot at sync, mutations queue at sync. */
export function createStrictExcelDouble(seed = {}, faults = {}) {
  const state = { sheets: [...(seed.sheets ?? ["Sheet1"])], cells: clone(seed.cells ?? {}), formulas: clone(seed.formulas ?? {}), formats: {}, numberFormats: {}, tables: [], comments: [] };
  let mutationCount = 0; let mutationGroups = 0; let runs = 0; let apiErrors = 0;
  const owners = new WeakMap();
  const rangeLocal = new WeakMap();
  const Excel = { async run(callback) {
    runs++;
    if (faults.staleBeforeExecute && runs === 2) state.cells[`${state.sheets[0]}!A1`] = "stale-synthetic";
    let pendingLoads = []; let mutations = []; let active = true; const owner = {};
    const requireActive = () => { if (!active) { apiErrors++; throw coded("ContextExpired"); } };
    const queue = action => { requireActive(); mutations.push(action); };
    function proxy(readers, methods = {}, children = {}, setters = {}, itemFields = undefined) {
      const loaded = new Map();
      const snapshotLoad = field => {
        requireActive();
        const [first, ...rest] = field.split("/");
        if (!Object.hasOwn(readers, first)) { apiErrors++; throw coded("ApiNotFound"); }
        // Only collection (items) descendants may be loaded; any other nested
        // path is an unsupported API shape and must not silently succeed.
        if (rest.length && first !== "items") { apiErrors++; throw coded("ApiNotFound"); }
        // Validate the entire descendant path against the item schema (items are
        // scalar fields), so empty collections cannot hide a bad path.
        if (rest.length && first === "items" && (!itemFields || rest.length !== 1 || !itemFields.has(rest[0]))) { apiErrors++; throw coded("ApiNotFound"); }
        // Snapshot the value at sync time; collection descendants are drained
        // in the same sync so their loads resolve before the sync returns.
        const value = readers[first]();
        loaded.set(first, Array.isArray(value) && first !== "items" ? clone(value) : value);
        if (rest.length && first === "items") for (const item of value) item.load(rest.join("/"));
      };
      const guardedMethods = Object.fromEntries(Object.entries(methods).map(([key, method]) => [key, (...args) => { requireActive(); return method(...args); }]));
      const api = { ...guardedMethods, load(fields) {
        requireActive();
        // Reject invalid load arguments (null, numbers, non-string array items)
        // with a recorded API error before any split can throw an unaccounted TypeError.
        if (typeof fields !== "string" && !(Array.isArray(fields) && fields.every(f => typeof f === "string"))) { apiErrors++; throw coded("ApiNotFound"); }
        for (const field of (Array.isArray(fields) ? fields : fields.split(","))) pendingLoads.push(() => snapshotLoad(field.trim()));
      } };
      const object = new Proxy(api, {
        get(target, key) {
          if (typeof key === "symbol" || key === "then") return undefined;
          requireActive();
          if (Object.hasOwn(target, key)) return target[key];
          if (Object.hasOwn(children, key)) return children[key];
          if (Object.hasOwn(readers, key)) {
            if (!loaded.has(key)) { apiErrors++; throw coded("PropertyNotLoaded"); }
            return loaded.get(key);
          }
          apiErrors++; throw coded("ApiNotFound");
        },
        set(_target, key, value) {
          requireActive();
          if (!Object.hasOwn(setters, key)) { apiErrors++; throw coded("ApiNotFound"); }
          // Record clone failures (non-cloneable setter inputs) as API errors
          // instead of letting a swallowed DataCloneError evade the sticky error count.
          let snapshot;
          try { snapshot = clone(value); } catch { apiErrors++; throw coded("ApiNotFound"); }
          queue(() => setters[key](snapshot)); return true;
        },
      });
      owners.set(object, owner); return object;
    }
    function range(sheet, address, nullObject = false) {
      let b; try { b = bounds(address); } catch (error) { apiErrors++; throw error; }
      const keys = Array.from({ length: b.rows }, (_, row) => Array.from({ length: b.cols }, (_, col) => `${sheet}!${column(b.start.col + col)}${b.start.row + row}`));
      const matrix = (map, fallback) => keys.map(row => row.map(k => map[k] ?? fallback));
      const assign = (map, values) => { if (!Array.isArray(values) || values.length !== b.rows || values.some(row => !Array.isArray(row) || row.length !== b.cols || row.some(v => !(v === null || typeof v === "string" || typeof v === "number" || typeof v === "boolean")))) { apiErrors++; throw coded("double_shape_mismatch"); } if (!state.sheets.includes(sheet)) { apiErrors++; throw coded("ItemNotFound"); } keys.forEach((row, i) => row.forEach((key, j) => { map[key] = values[i][j]; })); };
      const formatState = state.formats[`${sheet}!${b.local}`] ??= { font: {}, fill: {} };
      const nestedFormat = (kind, props) => proxy(Object.fromEntries(props.map(k => [k, () => formatState[kind][k]])), {}, {}, Object.fromEntries(props.map(k => [k, v => { if (!state.sheets.includes(sheet)) { apiErrors++; throw coded("ItemNotFound"); } formatState[kind][k] = v; }])));
      const format = proxy({}, {}, { font: nestedFormat("font", ["bold", "italic", "color", "size", "name", "underline"]), fill: nestedFormat("fill", ["color"]) });
      const requireSheet = () => { if (!state.sheets.includes(sheet)) { apiErrors++; throw coded("ItemNotFound"); } };
      const object = proxy({ address: () => `${sheet}!${b.local}`, rowCount: () => b.rows, columnCount: () => b.cols, isNullObject: () => nullObject,
        values: () => { requireSheet(); return matrix(state.cells, null); }, formulas: () => { requireSheet(); return matrix(state.formulas, ""); }, numberFormat: () => { requireSheet(); return matrix(state.numberFormats, "General"); } }, { select: () => { requireSheet(); } }, { format }, {
        values: v => assign(state.cells, v), formulas: v => {
          assign(state.formulas, v);
          // Primitive write_range assigns formulas even for literal values.
          // Reflect literals in values; formula calculation is not simulated.
          keys.forEach((row, i) => row.forEach((key, j) => {
            if (typeof v[i][j] !== "string" || !v[i][j].startsWith("=")) state.cells[key] = v[i][j];
          }));
        }, numberFormat: v => assign(state.numberFormats, v),
      });
      rangeLocal.set(object, { sheet, local: b.local }); return object;
    }
    function tableObject(record, name) {
      const get = () => record ?? state.tables.find(t => t.name === name);
      return proxy({ name: () => get()?.name ?? name, isNullObject: () => !get(), showHeaders: () => get()?.showHeaders, showTotals: () => get()?.showTotals }, {
        getRange: () => { const t = get(); if (!t) { apiErrors++; throw coded("ItemNotFound"); } return range(t.sheet, t.address); },
      }, { columns: proxy({ items: () => [] }, {}, {}, {}, new Set()) }, { name: value => { if (!get()) { apiErrors++; throw coded("missing_table"); } get().name = value; } });
    }
    const tableCollection = proxy({ items: () => state.tables.map(t => tableObject(t)) }, {
      getItem: name => { if (!state.tables.some(t => t.name === name)) { apiErrors++; throw coded("missing_table"); } return tableObject(undefined, name); },
      getItemOrNullObject: name => tableObject(undefined, name),
    }, {}, {}, new Set(["name", "isNullObject", "showHeaders", "showTotals"]));
    function sheetObject(name) {
      const tables = proxy({}, { add: (r, hasHeaders) => {
        requireActive();
        if (owners.get(r) !== owner) { apiErrors++; throw coded("InvalidRequestContext"); }
        // Reject non-range objects even from the same context, so an adapter
        // argument-type regression cannot fabricate a table from any proxy.
        if (!rangeLocal.has(r)) { apiErrors++; throw coded("ApiNotFound"); }
        if (!state.sheets.includes(name)) { apiErrors++; throw coded("ItemNotFound"); }
        const rangeIdentity = rangeLocal.get(r);
        if (rangeIdentity.sheet !== name) { apiErrors++; throw coded("InvalidRequestContext"); }
        const record = { name: `Table${state.tables.length + 1}`, sheet: name, address: rangeIdentity.local, showHeaders: hasHeaders, showTotals: false };
        queue(() => { state.tables.push(record); });
        return tableObject(record);
      } });
      const comments = proxy({ items: () => [] }, { getCount: () => { const result = proxy({ value: () => 0 }); result.load("value"); return result; } }, {}, {}, new Set());
      return proxy({ name: () => name, isNullObject: () => !state.sheets.includes(name) }, {
        getRange: address => range(name, address),
        getUsedRangeOrNullObject: () => range(name, "A1:D10", false),
      }, { tables, comments });
    }
    const worksheets = proxy({ items: () => state.sheets.map(sheetObject) }, {
      getItem: name => { if (!state.sheets.includes(name)) { apiErrors++; throw coded("missing_sheet"); } return sheetObject(name); },
      getItemOrNullObject: sheetObject,
      getActiveWorksheet: () => sheetObject(state.sheets[0]),
      add: name => { queue(() => state.sheets.push(name)); return sheetObject(name); },
    }, {}, {}, new Set(["name", "isNullObject"]));
    const workbook = proxy({}, { getSelectedRange: () => range(state.sheets[0], seed.selection ?? "A1:D10") }, { worksheets, tables: tableCollection });
    const context = proxy({}, { async sync() {
      requireActive();
      // Apply queued writes first so a same-sync load snapshots the post-write
      // state, then snapshot deferred loads. Snapshots are frozen, so a range
      // written through another proxy still requires its own reload.
      if (mutations.length) {
        mutationGroups++;
        if (faults.failMutationGroup === mutationGroups) { mutations = []; throw coded("scripted_sync_failure"); }
        const batch = mutations; mutations = []; for (const mutation of batch) { mutation(); mutationCount++; }
        if (faults.divergeAfterMutation) state.cells[faults.divergeAfterMutation.cell] = faults.divergeAfterMutation.value;
      }
      while (pendingLoads.length) { const loads = pendingLoads; pendingLoads = []; for (const load of loads) load(); }
    } }, { workbook });
    try { return await callback(context); }
    finally { active = false; pendingLoads = []; mutations = []; }
  } };
  return { Excel, get mutationCount() { return mutationCount; }, get apiErrors() { return apiErrors; }, get runCount() { return runs; },
    get snapshot() { return { cells: clone(state.cells), numberFormats: clone(state.numberFormats), tables: state.tables.map(t => ({ ...t })) }; } };
}
function response(output) {
  return typeof output === "string" ? { content: output, toolCalls: [] } : { content: output?.content ?? "", toolCalls: output?.tool_calls ?? output?.toolCalls ?? [] };
}
function hasPlan(output) {
  return typeof output === "string" ? output.includes("excel_execute_plan") : (output?.tool_calls ?? output?.toolCalls ?? []).some(c => c.function?.name === "excel_execute_plan");
}
export async function checkFixtureProtocol(kernel, fixture) {
  const r = response(fixture.modelOutput);
  const normalized = await kernel.normalizeWorkbookPlanResponse(r.content, JSON.stringify(r.toolCalls));
  return { caseId: fixture.caseId, accepted: normalized.ok === true && normalized.kind === "plan", reasonCode: normalized.ok === false ? "protocol_rejected" : normalized.kind };
}
async function evaluateServerCase(kernel, fixture, outputs, double, inspect, clock) {
  // R4g harness pattern: real coordinator/bindings/route, fake companion delivery and ACK.
  const bindings = new kernel.ServerOwnedTurnBindingStore({ now: clock, ttlMs: 90000, maxEntries: 32, maxBytes: 65536 });
  const coordinator = new kernel.TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => ids.documentId,
    getBridgeIdentity: () => ({ sessionId: ids.sessionId, documentId: ids.documentId }),
    reconnectForDocumentChange: async () => { throw coded("unexpected_document_change"); }, bindings, paneId: "66666666-6666-4666-8666-666666666666",
  });
  const controller = new AbortController(); const turnStart = Date.now();
  const trustedRequest = kernel.classifyWorkbookRequest(fixture.syntheticInput.userPrompt);
  const handle = coordinator.beginTurn(ids.turnId, controller, { trustedRequest, turnStart, cancellationSignal: controller.signal, workbookPlanRolloutEnabled: true });
  if ((await coordinator.prepareTurn(handle)).status !== "active") throw coded("server_turn_not_active");
  const store = new kernel.BrowserIdempotencyStore(new kernel.MemoryIdempotencyStorage());
  const route = kernel.createServerOwnedPlanRoute({
    isRolloutEnabled: () => true, getBridgeSessionId: () => ids.sessionId,
    getActiveTurn: () => ({ turnOperationId: ids.turnId, trustedRequest, sessionId: ids.sessionId, turnStart, hasCancellation: true }),
    coordinator, bindings, now: clock,
    readRuntimeContext: plan => kernel.readNativeWorkbookPlanRuntimeContext(plan),
    createAdapter: () => kernel.createNativeWorkbookPlanAdapter(), createStore: () => store,
    inspect: async reads => {
      // Track actual inspection invocation and requested reads, so a route that
      // returns "inspected" without running inspection cannot pass.
      inspectionReadCount++; inspectionReads = reads?.length ?? 0;
      return JSON.stringify({ ok: true, reads: JSON.parse(await inspect(reads)).map(data => ({ ok: true, data })) });
    },
    awaitPaneConfirmation: (binding, signal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal),
    executePrimitive: async () => { throw coded("unexpected_primitive"); },
  });
  let body; let ackCount = 0; let opCount = 0; let inspectionCount = 0; let inspectionReadCount = 0; let inspectionReads = 0;
  for (let i = 0; i < outputs.length; i++) {
    const normalized = await kernel.normalizeWorkbookPlanResponse(response(outputs[i]).content, JSON.stringify(response(outputs[i]).toolCalls));
    if (normalized.ok !== true || normalized.kind !== "plan") throw coded("server_fixture_not_plan");
    const plan = normalized.plan; opCount = plan.operations?.length ?? 0;
    body = JSON.parse(await route.executeCompanionTool("excel_execute_plan", plan, {
      turnOperationId: i === 0 ? "44444444-4444-4444-8444-444444444444" : "55555555-5555-4555-8555-555555555555",
      workbookPlanFingerprint: kernel.workbookPlanFingerprint, phase: plan.outcome === "inspect" ? "inspect" : "execute", turnKey: ids.turnId,
    }, { acknowledgeExecuteStart: async () => { if (double.mutationCount !== 0) throw coded("ack_after_mutation"); ackCount++; } }));
    if (plan.outcome === "inspect") {
      if (body.result?.outcome !== "inspected" || double.mutationCount !== 0 || ackCount !== 0) throw coded("server_inspection_failed");
      if (inspectionReadCount !== inspectionCount + 1 || inspectionReads < 1) throw coded("server_inspection_failed");
      inspectionCount++;
    }
  }
  if (inspectionCount !== 1 || ackCount !== 1 || inspectionReadCount !== 1) throw coded("server_phase_or_ack_failed");
  const result = body.result;
  return { route: "server-owned", stageReached: "verification", outcome: result?.outcome ?? body.reason,
    verification: result?.verification ?? "not_attempted", rollbackStatus: result?.execution?.rollback ?? "not_attempted", mutationCount: double.mutationCount, opCount, state: double.snapshot };
}
async function evaluateCase(kernel, fixture, clock, frozen = false) {
  const outputs = Array.isArray(fixture.modelOutput) ? fixture.modelOutput : [fixture.modelOutput];
  const isPlan = outputs.some(hasPlan);
  if (!fixture.provenance && !isPlan) {
    const r = response(outputs[0]); await kernel.normalizeWorkbookPlanResponse(r.content, JSON.stringify(r.toolCalls));
    return { route: "primitive_not_plan", stageReached: "protocol", outcome: "not_plan_candidate", verification: "not_attempted", rollbackStatus: "not_attempted", mutationCount: 0, opCount: 0, status: "unavailable", reasonCode: "primitive_not_plan", protocolAccepted: null };
  }
  let protocolAccepted = null;
  if (frozen) {
    try { protocolAccepted = (await checkFixtureProtocol(kernel, fixture)).accepted; }
    catch { return { route: "direct", stageReached: "protocol", outcome: "error", verification: "not_attempted", rollbackStatus: "not_attempted", mutationCount: 0, opCount: 0, status: "error", reasonCode: "protocol_probe_error", protocolAccepted }; }
  }
  const seed = fixture.workbookSeed ?? { sheets: [fixture.syntheticInput.workbookContext.activeSheet], cells: {} };
  const double = createStrictExcelDouble(seed, fixture.faults);
  const previous = Object.getOwnPropertyDescriptor(globalThis, "Excel");
  let runtime; let stageReached = "protocol"; let index = 0; let operationCount = 0;
  Object.defineProperty(globalThis, "Excel", { configurable: true, writable: true, value: double.Excel });
  try {
    const adapter = kernel.createNativeWorkbookPlanAdapter();
    const dependencies = {
      identity: ids, trustedOperationBoundary: kernel.createTrustedOperationBoundary(ids),
      model: async () => response(outputs[Math.min(index++, outputs.length - 1)]),
      inspect: async reads => { stageReached = "inspection"; return double.Excel.run(async c => {
        const targets = reads.map(read => { if (read.kind !== "read_range") throw coded("unsupported_inspection"); const r = c.workbook.worksheets.getItem(read.sheet).getRange(read.address); r.load("values"); return r; });
        await c.sync(); return JSON.stringify(targets.map(r => r.values));
      }); },
      getCompileContextForPlan: async plan => { stageReached = "runtime"; operationCount = plan.operations.length; runtime = await kernel.readNativeWorkbookPlanRuntimeContext(plan); return runtime.compileContext; },
      resourcePreflight: plan => { operationCount = plan.operations?.length ?? 0; stageReached = "resources"; return kernel.preflightWorkbookPlanResources(plan, { compileContext: { sheetNames: new Set(), tablesByName: new Map(), supportedPrimitiveNames: new Set(), supportedCapabilities: new Set(), chartSourceResolutionsBySheet: new Map() } }); },
      resourcePreflightAfterRuntime: plan => kernel.preflightWorkbookPlanResources(plan, runtime),
      policyInput: (plan, confirmation) => { stageReached = "policy_compile"; const trusted = kernel.classifyWorkbookRequest(fixture.syntheticInput.userPrompt); return kernel.buildWorkbookPlanPolicyInput(plan, runtime, { ...trusted, ...(confirmation ? { confirmation } : {}) }); },
      awaitPaneConfirmation: async () => fixture.confirmation,
      consumePaneConfirmation: (candidate, binding) => {
        if (candidate !== "confirm") return undefined;
        // A valid confirmation is still rejected once the turn deadline has
        // elapsed; fixtures opt into this via a controlled confirmation clock.
        const confirmationAtMs = Number.isFinite(fixture.confirmationAtMs) ? fixture.confirmationAtMs : clock();
        if (typeof binding?.deadline === "number" && confirmationAtMs >= binding.deadline) return undefined;
        return { eventId: "synthetic-confirmation", confirmed: true, source: "user" };
      },
      store: new kernel.BrowserIdempotencyStore(new kernel.MemoryIdempotencyStorage()), adapter, rolloutEnabled: true, route: "direct", now: clock,
    };
    const result = fixture.route === "server-owned" ? undefined : await kernel.createWorkbookPlanTurn(dependencies);
    const observed = fixture.route === "server-owned" ? await evaluateServerCase(kernel, fixture, outputs, double, dependencies.inspect, clock) : { route: "direct", stageReached: result.execution ? (result.verification === "not_attempted" ? "execution" : "verification") : stageReached,
      outcome: result.outcome, verification: result.verification, rollbackStatus: result.execution?.rollback ?? "not_attempted", mutationCount: double.mutationCount, opCount: result.execution?.operationCount ?? operationCount, state: double.snapshot };
    // Workbook state is an acceptance oracle only; evidence records stay privacy-redacted.
    const { state: _workbookState, ...publicRecord } = observed;
    if (double.apiErrors) return { ...publicRecord, status: "error", reasonCode: "strict_office_api_error", protocolAccepted };
    const expect = fixture.expect;
    // Frozen labels describe protocol acceptance, not full native success.
    const matches = frozen ? protocolAccepted === (fixture.expectedDisposition === "accepted") : expect && Object.entries(expect).every(([k, v]) => k === "state" ? JSON.stringify(canonical(observed[k])) === JSON.stringify(canonical(v)) : observed[k] === v);
    return { ...publicRecord, status: matches ? "pass" : "fail", reasonCode: matches ? "expected_disposition" : "expectation_mismatch", protocolAccepted };
  } catch (error) { return { route: fixture.route === "server-owned" ? "server-owned" : "direct", stageReached, outcome: "error", verification: "not_attempted", rollbackStatus: "not_attempted", mutationCount: double.mutationCount, opCount: operationCount, status: "error", reasonCode: error.code === "server_inspection_failed" || error.code === "server_phase_or_ack_failed" ? error.code : "offline_runtime_error", protocolAccepted }; }
  finally { if (previous) Object.defineProperty(globalThis, "Excel", previous); else delete globalThis.Excel; }
}
export async function runWorkbookPlanConformance({ kernel, corpora, clock = () => 1000, budgets, out, candidateSha = "unknown" }) {
  const candidate = validateCandidateSha(candidateSha);
  const destination = out === undefined ? undefined : validateOutputDirectory(out);
  const loaded = corpora.map(c => {
    if (typeof c === "string") return loadCorpus(c);
    if (c && typeof c === "object" && typeof c.path === "string") return loadCorpus(c.path, { role: c.role });
    if (c && typeof c === "object" && Array.isArray(c.fixtures)) return c;
    throw coded("invalid_corpus");
  });
  const hasInMemory = corpora.some(c => c && typeof c === "object" && typeof c.path !== "string" && Array.isArray(c.fixtures));
  for (const corpus of loaded) {
    validateManifestMetadata(corpus.manifest);
    if (!["original", "synthetic"].includes(corpus.role)) throw coded("invalid_corpus_role");
    if (corpus.role === "synthetic") for (const fixture of corpus.fixtures) assertPrivacyClean(fixture);
  }
  // In-memory corpora are a testing seam, not evidence: they carry no raw
  // bytes, so identity cannot be recomputed and must not produce records.
  if (destination !== undefined && hasInMemory) throw coded("inmemory_corpus_not_evidence", 2);
  const pins = loaded.filter(c => c.role === "synthetic").flatMap(c => c.manifest.contractDrift ?? []);
  const originals = loaded.filter(c => c.role === "original").flatMap(c => c.fixtures);
  const pinMap = new Map();
  for (const pin of pins) {
    const fixture = originals.find(f => f.caseId === pin.caseId);
    if (Object.keys(pin).sort().join(",") !== "caseId,provenanceCommit,reasonCode" || !fixture || !hasPlan(fixture.modelOutput) || pinMap.has(pin.caseId) || pin.reasonCode !== "legacy_plan_shape_never_accepted" || !/^[0-9a-f]{40}$/.test(pin.provenanceCommit)) throw coded("invalid_contract_drift_pin");
    pinMap.set(pin.caseId, pin);
  }
  const records = [];
  for (const { role, manifest, fixtures } of loaded) for (const fixture of fixtures) {
    const observed = await evaluateCase(kernel, fixture, clock, role === "original");
    const pin = role === "original" ? pinMap.get(fixture.caseId) : undefined;
    if (pin) {
      if (observed.status === "pass") Object.assign(observed, { status: "error", reasonCode: "contract_drift_pin_matches_label" });
      else if (observed.status === "fail" && observed.protocolAccepted === false && observed.outcome === "rejected" && observed.stageReached === "protocol" && observed.mutationCount === 0) Object.assign(observed, { status: "contractDrift", reasonCode: pin.reasonCode });
      else Object.assign(observed, { status: "error", reasonCode: "contract_drift_pin_outcome_changed" });
    }
    records.push({ evaluatorVersion: EVALUATOR_VERSION, candidateSha: candidate, schemaFingerprint: kernel.workbookPlanFingerprint,
      corpusId: manifest.corpusId ?? "model-conformance", corpusVersion: manifest.corpusVersion, corpusSha256: manifest.corpusSha256, caseId: fixture.caseId,
      transport: safeMetadata(fixture.transport, "unknown"), locale: safeMetadata(fixture.locale ?? "en", "unknown"), modelId: `synthetic:${safeMetadata(fixture.sourceClass, "generic")}`, modelRevision: "n/a",
      budgets: { turnTimeoutMs: kernel.WORKBOOK_PLAN_TURN_TIMEOUT_MS, maxCorrections: kernel.WORKBOOK_PLAN_MAX_CORRECTIONS, resourceLimits: kernel.WORKBOOK_PLAN_RESOURCE_LIMITS }, context: { seedSha256: hash(JSON.stringify(fixture.workbookSeed ?? {})) }, ...observed });
  }
  void budgets; // Runtime budgets are owned by the real kernel, not an evaluator override.
  const summary = { total: records.length, pass: 0, fail: 0, unavailable: 0, error: 0, contractDrift: 0, notPlanCandidate: 0 };
  for (const r of records) { summary[r.status]++; if (r.route === "primitive_not_plan") summary.notPlanCandidate++; }
  if (summary.total !== summary.pass + summary.fail + summary.unavailable + summary.error + summary.contractDrift) throw coded("summary_mismatch");
  if (destination) { mkdirSync(destination, { recursive: true }); writeFileSync(join(destination, "records.jsonl"), [...records, summary].map(r => JSON.stringify(r)).join("\n") + "\n"); }
  return { records, summary };
}
/** Synthetic primary executes through the real primitive authority; shadow is
 * pure protocol/validation/policy/compile only. No production shadow route. */
export async function runWorkbookPlanShadowComparison({ kernel, corpus, out, candidateSha = "unknown", clock }) {
  const candidate = validateCandidateSha(candidateSha);
  const destination = out === undefined ? undefined : validateOutputDirectory(out);
  const loaded = loadCorpus(corpus, { role: "synthetic" });
  if (loaded.manifest.corpusId !== "workbook-plan-shadow") throw coded("invalid_shadow_corpus");
  // Targets are hashed with a per-run salt in memory; only equality/counts leave.
  const salt = randomBytes(32).toString("hex");
  const targetHashes = primitives => [...new Set(primitives.map(p => hash(salt + JSON.stringify([p.args.sheet, p.args.address]))))].sort();
  const equal = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const effects = primitives => [...new Set(primitives.map(p => p.effect))].sort();
  const records = [];
  let tick = 0; const now = clock ?? (() => ++tick);
  for (const fixture of loaded.fixtures) {
    const previous = Object.getOwnPropertyDescriptor(globalThis, "Excel");
    const double = createStrictExcelDouble(fixture.workbookSeed);
    const trusted = kernel.classifyWorkbookRequest(fixture.syntheticInput.userPrompt);
    const primary = []; let primaryOutcome = "committed";
    const shadow = []; let shadowOutcome = "rejected"; let shadowStage = "protocol"; let shadowExecuteCalls = 0;
    let primaryTicks = 0; let shadowTicks = 0; let failed = false;
    const start = now();
    try {
      Object.defineProperty(globalThis, "Excel", { configurable: true, writable: true, value: double.Excel });
      const authority = kernel.createPrimitiveMutationAuthority({ trustedRequest: trusted, workbookPlanRolloutEnabled: true });
      try {
        for (const call of fixture.primary) {
          const result = JSON.parse(await authority.executePrimitive(call.name, call.args));
          if (result.ok !== true) { primaryOutcome = "denied"; break; }
          const effect = kernel.getPrimitiveEffect(call.name);
          primary.push({ args: call.args, effect: effect?.planEffect ?? effect?.effectClass });
        }
      } finally { authority.terminate(); }
      if (double.apiErrors) throw coded("strict_office_api_error");
      primaryTicks = now() - start;
      const shadowStart = now();
      Object.defineProperty(globalThis, "Excel", { configurable: true, writable: true, value: {
        run() { shadowExecuteCalls++; throw coded("shadow_execution_forbidden"); },
      } });
      const r = response(fixture.modelOutput);
      const normalized = await kernel.normalizeWorkbookPlanResponse(r.content, JSON.stringify(r.toolCalls));
      if (normalized.ok === true && normalized.kind === "plan") {
        shadowStage = "validation";
        const validated = kernel.validateWorkbookPlan(normalized.plan);
        if (validated.ok === true && normalized.plan.outcome === "execute") {
          const plan = normalized.plan;
          // Synthetic, fixture-owned pre-state only; never read Office on shadow.
          // No precondition manifest is fabricated and no executable transaction
          // authority is minted from this pure compiler inspection shape.
          const compileContext = {
            sheetNames: new Set(fixture.workbookSeed.sheets), tablesByName: new Map(),
            supportedPrimitiveNames: new Set(kernel.workbookPlanContract.operationDefinitions.flatMap(d => d.compilerTargets)),
            supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
            chartSourceResolutionsBySheet: new Map(),
          };
          for (const name of fixture.compileDisabledPrimitives ?? []) compileContext.supportedPrimitiveNames.delete(name);
          const runtime = { compileContext, activeSheetName: fixture.syntheticInput.workbookContext.activeSheet,
            selection: { sheetName: fixture.syntheticInput.workbookContext.activeSheet, address: fixture.syntheticInput.workbookContext.selectedRange } };
          shadowStage = "policy";
          const decision = kernel.decideWorkbookPlanPolicy(kernel.createWorkbookPlanPolicyInputSnapshot(kernel.buildWorkbookPlanPolicyInput(plan, runtime, trusted)));
          shadowOutcome = decision.decision;
          if (decision.decision === "approved") {
            shadowStage = "compile";
            const compiled = kernel.compileWorkbookPlan(plan, compileContext);
            shadowOutcome = compiled.ok === true ? "compiled" : "rejected";
            if (compiled.ok === true) for (const p of compiled.transaction.primitives) shadow.push({ args: p.args, effect: p.effect });
          }
        }
      }
      shadowTicks = now() - shadowStart;
      if (shadowExecuteCalls) throw coded("shadow_execution_forbidden");
    } catch {
      failed = true; // Fixed metadata only, never export host errors or arguments.
    } finally {
      if (previous) Object.defineProperty(globalThis, "Excel", previous); else delete globalThis.Excel;
    }
    const a = targetHashes(primary); const b = targetHashes(shadow);
    const observed = { primaryOutcome, shadowStage, shadowOutcome, targetSetEqual: equal(a, b), primaryTargetCount: a.length, shadowTargetCount: b.length,
      effectClassesEqual: equal(effects(primary), effects(shadow)), primaryOpCount: primary.length, shadowOpCount: shadow.length,
      latencyTicks: { primary: primaryTicks, shadow: shadowTicks }, tokens: fixture.usage ?? null, shadowExecuteCalls };
    const matches = Object.entries(fixture.expect).every(([k, v]) => observed[k] === v);
    records.push({ recordType: "shadow", evaluatorVersion: EVALUATOR_VERSION, candidateSha: candidate,
      schemaFingerprint: kernel.workbookPlanFingerprint, corpusSha256: loaded.manifest.corpusSha256, caseId: fixture.caseId,
      ...observed, status: failed ? "error" : matches ? "pass" : "fail" });
  }
  const summary = { total: records.length, pass: 0, fail: 0, error: 0, mismatches: 0 };
  for (const r of records) { summary[r.status]++; if (!r.targetSetEqual || !r.effectClassesEqual || r.primaryOpCount !== r.shadowOpCount) summary.mismatches++; }
  if (summary.total !== summary.pass + summary.fail + summary.error) throw coded("summary_mismatch");
  if (destination) {
    // validateOutputDirectory covers records.jsonl; reject a pre-existing shadow
    // leaf too (including symlinks), then use exclusive creation for this leaf.
    const leaf = join(destination, "shadow-records.jsonl");
    if (lstatSync(leaf, { throwIfNoEntry: false })) throw coded("output_alias_forbidden", 2);
    mkdirSync(destination, { recursive: true });
    writeFileSync(leaf, [...records, summary].map(r => JSON.stringify(r)).join("\n") + "\n", { flag: "wx" });
  }
  return { records, summary };
}
export function loadCompiledKernel() {
  const scratch = validateScratchDirectory(process.env.TMPDIR);
  // Retain the offline compiled loader in scratch: no deletion or build writes in the repo.
  const build = mkdtempSync(join(scratch, "r6b-conformance-build-"));
  symlinkSync(join(repoRoot, "node_modules"), join(build, "node_modules"), "dir");
  const modules = ["workbookPlanOrchestrator", "workbookPlanOfficeAdapter", "workbookPlanExecutor", "workbookPlanRequest", "workbookPlanRoute", "workbookPlanProtocol", "workbookPlanResources", "generated/workbookPlanSchema", "taskpane/serverOwnedPlanRoute", "taskpane/taskPaneTurnCoordinator", "primitiveMutationAuthority", "excelTools", "workbookPlanCompiler", "workbookPlanValidation", "workbookPlanPolicy", "workbookPlanContract"];
  execFileSync(process.execPath, [join(repoRoot, "node_modules/typescript/bin/tsc"), ...modules.map(m => `src/${m}.ts`), "--target", "ES2022", "--module", "commonjs", "--moduleResolution", "node", "--esModuleInterop", "--skipLibCheck", "--types", "office-js", "--outDir", build], { cwd: repoRoot, stdio: "pipe" });
  const require = createRequire(import.meta.url);
  return Object.assign({}, ...modules.map(m => require(join(build, `${m}.js`))));
}
export async function main(args = process.argv.slice(2)) {
  try {
    const allowed = new Set(["--out", "--candidate", "--original-corpus", "--corpus", "--shadow", "--shadow-corpus"]); const flags = {};
    for (let i = 0; i < args.length; i++) {
      const key = args[i];
      if (!allowed.has(key) || Object.hasOwn(flags, key)) throw coded("invalid_arguments", 2);
      if (key === "--shadow") { flags[key] = true; continue; }
      if (!args[i + 1] || args[i + 1].startsWith("--")) throw coded("invalid_arguments", 2);
      flags[key] = args[++i];
    }
    const out = validateOutputDirectory(flags["--out"]);
    if (flags["--shadow"]) {
      const result = await runWorkbookPlanShadowComparison({ kernel: loadCompiledKernel(), corpus: flags["--shadow-corpus"] ?? join(repoRoot, "src/__tests__/fixtures/workbook-plan-conformance/shadow-v1"), out, candidateSha: flags["--candidate"] });
      console.log(JSON.stringify(result.summary)); return result.summary.fail || result.summary.error ? 1 : 0;
    }
    const corpora = [{ path: flags["--original-corpus"] ?? join(repoRoot, "src/__tests__/fixtures/model-conformance"), role: "original" }, { path: flags["--corpus"] ?? join(repoRoot, "src/__tests__/fixtures/workbook-plan-conformance/v1"), role: "synthetic" }];
    const result = await runWorkbookPlanConformance({ kernel: loadCompiledKernel(), corpora, out, candidateSha: flags["--candidate"] });
    console.log(JSON.stringify(result.summary)); return result.summary.fail || result.summary.error ? 1 : 0;
  } catch (error) { console.error(JSON.stringify({ reasonCode: error.code ?? "evaluator_error" })); return error.exitCode ?? 1; }
}
if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) process.exitCode = await main();
