import { WORKBOOK_PLAN_FINGERPRINT, type WorkbookPlan, type InspectionOperation } from "../generated/workbookPlanSchema";
import { createWorkbookPlanTurn, ServerOwnedTurnBindingStore, type WorkbookPlanTurnResult, type WorkbookPlanTurnDependencies } from "../workbookPlanOrchestrator";
import { BrowserIdempotencyStore, MemoryIdempotencyStorage, createTrustedOperationBoundary, isReplayedWorkbookPlanReport } from "../workbookPlanExecutor";
import { createNativeWorkbookPlanAdapter, readNativeWorkbookPlanRuntimeContext } from "../workbookPlanOfficeAdapter";
import { buildWorkbookPlanPolicyInput, preflightWorkbookPlanResources, type WorkbookPlanRouteRuntimeContext } from "../workbookPlanRoute";
import { classifyWorkbookRequest } from "../workbookPlanRequest";
import { WORKBOOK_PLAN_TOOL_NAME } from "../workbookPlanProtocol";
import { createServerOwnedPlanRoute, EMPTY_RESOURCE_CONTEXT, type ServerOwnedPlanNotExecutedReason } from "./serverOwnedPlanRoute";
import { TaskPaneTurnCoordinator, type TaskPaneTurnHandle } from "./taskPaneTurnCoordinator";

declare const __HERMES_OFFICE_ACCEPTANCE__: boolean;
const compiledAcceptanceEnabled = () => typeof __HERMES_OFFICE_ACCEPTANCE__ !== "undefined" && __HERMES_OFFICE_ACCEPTANCE__ === true;

/** A future native launcher MUST create a new workbook, never reuse the active user workbook.
 * Identity is host-owned and checked before every native mutation handoff. No native
 * workbook creation, registration or live launch is performed by importing this module. */
export interface WorkbookPlanAcceptanceWorkbook {
  createDisposableWorkbook(runId: string, sheetNames: readonly string[]): Promise<void>;
  assertDisposableIdentity(runId: string, sheetNames: readonly string[]): Promise<void>;
  cleanupDisposableWorkbook(runId: string, sheetNames: readonly string[]): Promise<void>;
}
export type WorkbookPlanAcceptanceCase = "execute" | "inspect-execute" | "ask" | "refuse" | "invalid" | "replay" | "stop" | "rotation" | "caps" | "rollout-off";
export interface WorkbookPlanAcceptanceReceipt {
  readonly route: "direct" | "server-owned";
  readonly caseId: WorkbookPlanAcceptanceCase;
  readonly outcome: string;
  readonly executionOutcome: string;
  readonly reason: ServerOwnedPlanNotExecutedReason | "none";
  readonly verification: string;
  readonly inspectionCount: number;
  readonly correctionCount: number;
  readonly pass: boolean;
}
const CASES: readonly WorkbookPlanAcceptanceCase[] = ["execute", "inspect-execute", "ask", "refuse", "invalid", "replay", "stop", "rotation", "caps", "rollout-off"];
const REASONS: readonly string[] = ["ask", "refuse", "rejected", "cancelled", "not_started", "invalid", "already_executed"];
const uuid = (n: number) => `77777777-7777-4777-8777-${String(n).padStart(12, "0")}`;

export function isWorkbookPlanAcceptanceLocationAllowed(location: { href: string }): boolean {
  try {
    const url = new URL(location.href);
    const keys = [...url.searchParams.keys()];
    return url.origin === "https://localhost:3000" && url.pathname === "/taskpane.html" && !url.hash &&
      !url.username && !url.password && url.searchParams.getAll("workbookPlanAcceptance").length === 1 &&
      url.searchParams.get("workbookPlanAcceptance") === "1" && url.searchParams.getAll("_host_Info").length <= 1 &&
      keys.every(key => key === "workbookPlanAcceptance" || key === "_host_Info");
  } catch { return false; }
}

/** Only the separate acceptance entry may expose this API. Rollout is a local
 * turn dependency; settings, private state and the production pane are never changed. */
export function installWorkbookPlanAcceptanceHarness(
  target: Record<string, unknown>, workbook: WorkbookPlanAcceptanceWorkbook, location: { href: string },
): boolean {
  delete target.__hermesExcelWorkbookPlanAcceptance;
  if (!compiledAcceptanceEnabled() || !isWorkbookPlanAcceptanceLocationAllowed(location)) return false;
  target.__hermesExcelWorkbookPlanAcceptance = Object.freeze({ run: (runId: string) => runWorkbookPlanAcceptance(workbook, runId) });
  return true;
}

export async function runWorkbookPlanAcceptance(
  workbook: WorkbookPlanAcceptanceWorkbook, runId = "acceptance01",
): Promise<readonly WorkbookPlanAcceptanceReceipt[]> {
  if (!compiledAcceptanceEnabled()) throw new Error("WorkbookPlan harness requires the acceptance build");
  if (!/^[a-z0-9]{8,16}$/.test(runId)) throw new Error("Invalid synthetic run ID");
  const sheetNames = Array.from({ length: 7 }, (_, i) => `R7H_${runId}_${i}`);
  const sheet = sheetNames[0];
  const assertIdentity = () => workbook.assertDisposableIdentity(runId, sheetNames);
  const receipts: WorkbookPlanAcceptanceReceipt[] = [];
  // finally also covers partially failed creation. Cleanup failure rejects the run.
  try {
    await workbook.createDisposableWorkbook(runId, sheetNames);
    await assertIdentity();
    for (const route of ["direct", "server-owned"] as const) {
      for (const caseId of CASES) {
        await assertIdentity();
        const store = new BrowserIdempotencyStore(new MemoryIdempotencyStorage());
        let sessionId = uuid(1);
        const documentId = uuid(2);
        let bindingClockOffset = 0;
        const bindings = new ServerOwnedTurnBindingStore({ now: () => Date.now() + bindingClockOffset, ttlMs: 90_000, maxEntries: 32, maxBytes: 64 * 1024 });
        let rotations = 0;
        const coordinator = new TaskPaneTurnCoordinator({
          getOrCreateDocumentId: async () => documentId,
          getBridgeIdentity: () => ({ sessionId, documentId }),
          reconnectForDocumentChange: async () => { rotations++; sessionId = uuid(99); return { sessionId, documentId }; },
          bindings, paneId: uuid(3),
        });
        const trustedRequest = classifyWorkbookRequest("write synthetic-value into A1");
        let handle: TaskPaneTurnHandle;
        let turnStart = Date.now();
        async function begin(turn: number) {
          const controller = new AbortController(); turnStart = Date.now();
          handle = coordinator.beginTurn(uuid(turn), controller, { trustedRequest, turnStart, cancellationSignal: controller.signal, workbookPlanRolloutEnabled: true });
          const prepared = await coordinator.prepareTurn(handle);
          if (prepared.status !== "active") throw new Error("Acceptance turn identity unavailable");
        }
        await begin(4);
        const execute: WorkbookPlan = {
          version: 1, outcome: "execute", summary: "synthetic write",
          operations: [{ kind: "write_values", sheet, address: "A1", values: [["synthetic-value"]] }],
          verification: [{ kind: "range_values_equal", sheet, address: "A1", expectedValues: [["synthetic-value"]] }],
        };
        const inspect: WorkbookPlan = { version: 1, outcome: "inspect", reason: "synthetic inspection", reads: [{ kind: "read_range", sheet, address: "A1" }] };
        const plan: unknown = caseId === "ask" || caseId === "refuse" ? { version: 1, outcome: caseId, message: "synthetic response" }
          : caseId === "invalid" ? { version: 999, outcome: "execute" }
          : caseId === "caps" ? { ...execute, operations: Array.from({ length: 31 }, () => execute.operations[0]) }
          : execute;
        const native = createNativeWorkbookPlanAdapter();
        const adapter: WorkbookPlanTurnDependencies["adapter"] = {
          ...native,
          executeFirstGroup: async (gate, signal) => { await assertIdentity(); return native.executeFirstGroup!(gate, signal); },
          executeGroup: async (group, index, signal) => { await assertIdentity(); return native.executeGroup(group, index, signal); },
        };
        let inspectionCalls = 0;
        async function inspectNative(reads: readonly InspectionOperation[]): Promise<string> {
          await assertIdentity();
          if (reads.length !== 1 || reads[0].kind !== "read_range" || reads[0].sheet !== sheet || reads[0].address !== "A1") throw new Error("Non-synthetic inspection");
          const excel = (globalThis as unknown as { Excel: { run: (cb: (ctx: { workbook: { worksheets: { getItem: (name: string) => { getRange: (address: string) => { load: (field: string) => void; values: unknown } } } }; sync: () => Promise<void> }) => Promise<string>) => Promise<string> } }).Excel;
          const result = await excel.run(async ctx => {
            const range = ctx.workbook.worksheets.getItem(sheet).getRange("A1"); range.load("values"); await ctx.sync();
            return JSON.stringify({ reads: [{ ok: true, data: { values: range.values } }] });
          });
          inspectionCalls++;
          if (caseId === "stop") coordinator.stopCurrentTurn();
          if (caseId === "rotation") {
            coordinator.beginHermesRequest(sessionId); // deliberately unclean, requires real rotation
            await begin(5); // aborts the old generation before its continuation
          }
          return result;
        }
        const response = (value: unknown) => ({ content: "", toolCalls: [{ id: uuid(6), type: "function" as const, function: { name: WORKBOOK_PLAN_TOOL_NAME, arguments: JSON.stringify(value) } }] });
        async function direct(value: unknown, inspectFirst = false): Promise<WorkbookPlanTurnResult> {
          let runtime: WorkbookPlanRouteRuntimeContext | undefined;
          let calls = 0;
          const ownedHandle = handle!;
          return createWorkbookPlanTurn({
            identity: { sessionId, documentId, turnId: ownedHandle.turnId },
            trustedOperationBoundary: createTrustedOperationBoundary({ sessionId, documentId, turnId: uuid(7) }),
            getCompileContextForPlan: async p => { await assertIdentity(); runtime = await readNativeWorkbookPlanRuntimeContext(p); return runtime.compileContext; },
            policyInput: p => { if (!runtime) throw new Error("Missing runtime"); return buildWorkbookPlanPolicyInput(p, runtime, trustedRequest); },
            resourcePreflight: p => preflightWorkbookPlanResources(p, EMPTY_RESOURCE_CONTEXT),
            resourcePreflightAfterRuntime: p => runtime ? preflightWorkbookPlanResources(p, runtime) : { ok: false, reason: "missing_runtime" },
            store, adapter, model: async () => response(inspectFirst && calls++ === 0 ? inspect : value), inspect: inspectNative,
            signal: ownedHandle.cancellationSignal, turnStart, rolloutEnabled: caseId !== "rollout-off", route: "direct",
          });
        }
        const server = createServerOwnedPlanRoute({
          isRolloutEnabled: () => caseId !== "rollout-off", getBridgeSessionId: () => sessionId,
          getActiveTurn: () => ({ turnOperationId: handle!.turnId, trustedRequest, sessionId, turnStart, hasCancellation: true }),
          coordinator, bindings, readRuntimeContext: async p => { await assertIdentity(); return readNativeWorkbookPlanRuntimeContext(p); },
          createAdapter: () => adapter, createStore: () => store, inspect: inspectNative,
          awaitPaneConfirmation: (binding, signal) => coordinator.waitForWorkbookPlanConfirmation(binding, signal),
          executePrimitive: async () => { throw new Error("Primitive fallback forbidden"); },
        });
        let acknowledgements = 0;
        const serverCall = async (value: unknown, operation: number, phase?: "inspect" | "execute") => JSON.parse(await server.executeServerOwnedWorkbookPlan(
          value as Record<string, unknown>, { turnOperationId: uuid(operation), workbookPlanFingerprint: WORKBOOK_PLAN_FINGERPRINT, ...(phase ? { phase, turnKey: uuid(8) } : {}) },
          { acknowledgeExecuteStart: async () => { await assertIdentity(); acknowledgements++; } },
        )) as { reason?: string; state?: string; result?: WorkbookPlanTurnResult; ok?: boolean };
        let result: WorkbookPlanTurnResult | undefined;
        let reason: WorkbookPlanAcceptanceReceipt["reason"] = "none";
        let outcome = "rejected";
        let seedCommitted = true;
        let staleCancelled = true;
        try {
          if (route === "direct") {
            if (caseId === "replay") seedCommitted = (await direct(execute)).outcome === "committed";
            result = await direct(plan, ["inspect-execute", "stop", "rotation"].includes(caseId));
            if (caseId === "rotation") { staleCancelled = result.outcome === "cancelled"; result = await direct(execute); }
            outcome = result.outcome;
            if (caseId === "replay" && result.execution !== undefined && isReplayedWorkbookPlanReport(result.execution)) reason = "already_executed";
          } else {
            if (["inspect-execute", "stop", "rotation"].includes(caseId)) {
              const inspected = await serverCall(inspect, 10, "inspect");
              if (caseId === "inspect-execute" && inspected.result?.outcome !== "inspected") seedCommitted = false;
              if (caseId === "stop") { reason = "cancelled"; outcome = inspected.reason ?? "rejected"; }
              if (caseId === "rotation") staleCancelled = inspected.reason === "cancelled" || inspected.reason === "rejected";
            }
            if (caseId === "replay") {
              seedCommitted = (await serverCall(execute, 11)).result?.outcome === "committed";
              bindingClockOffset += 90_001; // expire ticket tombstone only; browser execution receipt remains
              await begin(5); // fresh generation, same operation identity => store replay, not ticket replay
            }
            if (caseId !== "stop") {
              const payload = await serverCall(plan, 11, caseId === "inspect-execute" ? "execute" : undefined);
              result = payload.result;
              reason = REASONS.includes(payload.reason ?? "") ? payload.reason as ServerOwnedPlanNotExecutedReason : "none";
              outcome = result?.outcome ?? reason;
            }
          }
          let capsRejected = true;
          if (caseId === "caps") {
            const capPlans: WorkbookPlan[] = [
              { version: 1, outcome: "inspect", reason: "synthetic read cap", reads: [{ kind: "read_range", sheet, address: "A1:A5001" }] },
              { version: 1, outcome: "execute", summary: "synthetic write cap", operations: [{ kind: "write_formula", sheet, address: "A1:A5001", formula: "=1" }], verification: execute.verification },
            ];
            for (let i = 0; i < capPlans.length; i++) {
              await begin(20 + i);
              if (route === "direct") {
                const rejectedCap = await direct(capPlans[i]);
                capsRejected = capsRejected && rejectedCap.outcome === "rejected" && rejectedCap.correctionCount === 0 &&
                  rejectedCap.finalSummary === (i === 0 ? "Workbook inspection resource preflight failed." : "Workbook execution resource preflight failed.");
              }
              else capsRejected = capsRejected && (await serverCall(capPlans[i], 30 + i, i === 0 ? "inspect" : "execute")).reason === "rejected";
            }
          }
          const expected = caseId === "execute" || caseId === "inspect-execute" || caseId === "rotation" ? outcome === "committed" && result?.verification === "verified"
            : caseId === "replay" ? reason === "already_executed"
            : caseId === "ask" || caseId === "refuse" ? outcome === caseId
            : caseId === "stop" ? outcome === "cancelled"
            : caseId === "rollout-off" ? outcome === (route === "direct" ? "rejected" : "not_started")
            : outcome === "rejected" || outcome === "invalid";
          const ackExpected = route === "direct" ? true : acknowledgements === (["execute", "inspect-execute", "replay", "rotation"].includes(caseId) ? 1 : 0);
          receipts.push(Object.freeze({ route, caseId, outcome, reason, executionOutcome: result?.execution?.outcome ?? "none", verification: result?.verification ?? "not_attempted", inspectionCount: inspectionCalls, correctionCount: result?.correctionCount ?? 0,
            pass: expected && capsRejected && seedCommitted && staleCancelled && ackExpected && (caseId !== "rotation" || rotations === 1) && (caseId !== "inspect-execute" || inspectionCalls === 1) }));
        } finally { coordinator.stopCurrentTurn(); bindings.stop(); }
      }
    }
    return Object.freeze(receipts);
  } finally { await workbook.cleanupDisposableWorkbook(runId, sheetNames); }
}
