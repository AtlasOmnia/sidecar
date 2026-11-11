import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { classifyWorkbookRequest } from "../workbookPlanRequest";
import {
  createWorkbookPlanTurn,
  ServerOwnedTurnBindingStore,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanModelResponse,
  type WorkbookPlanTurnDependencies,
  type WorkbookPlanTurnIdentity,
  type WorkbookPlanVerificationAdapter,
} from "../workbookPlanOrchestrator";
import {
  BrowserIdempotencyStore,
  createTrustedOperationBoundary,
  MemoryIdempotencyStorage,
  type WorkbookPlanExecutionAdapter,
} from "../workbookPlanExecutor";
import type { VerifiedWorkbookCompileContext } from "../workbookPlanCompiler";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";
import type { WorkbookPlan } from "../generated/workbookPlanSchema";
import type { TrustedConfirmationEvent } from "../workbookPlanRequest";
import { TaskPaneTurnCoordinator } from "../taskpane/taskPaneTurnCoordinator";

const IDS: WorkbookPlanTurnIdentity = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};

const TASKPANE_HTML = readFileSync(resolve(__dirname, "../taskpane/taskpane.html"), "utf8");
const TASKPANE_SOURCE = readFileSync(resolve(__dirname, "../taskpane/taskpane.ts"), "utf8");
/** R2d-1 (M2/M3): server-owned route logic moved from taskpane.ts into this extracted module. */
function readRouteSource(): string {
  try {
    return readFileSync(resolve(__dirname, "../taskpane/serverOwnedPlanRoute.ts"), "utf8");
  } catch {
    return "";
  }
}
const ROUTE_SOURCE = readRouteSource();

function createCoordinator(now: () => number = () => Date.now()): TaskPaneTurnCoordinator {
  return new TaskPaneTurnCoordinator({
    getOrCreateDocumentId: async () => IDS.documentId,
    getBridgeIdentity: () => ({ sessionId: null, documentId: null }),
    reconnectForDocumentChange: async () => ({ sessionId: IDS.sessionId, documentId: IDS.documentId }),
    bindings: new ServerOwnedTurnBindingStore({ now: () => 0, ttlMs: 90_000, maxEntries: 8 }),
    paneId: IDS.sessionId,
    now,
  });
}

const EXECUTE_PLAN: Extract<WorkbookPlan, { readonly outcome: "execute" }> = {
  version: 1,
  outcome: "execute",
  summary: "opaque",
  operations: [{ kind: "write_values", sheet: "Sheet1", address: "A1", values: [["ok"]] }],
  verification: [{ kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["ok"]] }],
};

const COMPILE_CONTEXT: VerifiedWorkbookCompileContext = {
  sheetNames: new Set(["Sheet1"]),
  tablesByName: new Map(),
  supportedPrimitiveNames: new Set(["write_range"]),
  supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
  chartSourceResolutionsBySheet: new Map(),
};

const CONFIRMATION_REQUIRED_POLICY: WorkbookPlanPolicyInput = {
  requestClassification: "ambiguous_mutation",
  activeSheetName: "Sheet1",
  selectionScopedRequest: false,
  operationSummaries: [{
    operationIndex: 0,
    effect: "write",
    target: { operationIndex: 0, sheetNames: ["Sheet1"], selectionRelationship: "inside" },
  }],
  capabilitySummaries: ["excel_workbook_access", "workbook_plan_v1_operation_audit"],
  resourceUsage: {
    operationCount: 1,
    cellsRead: 1,
    cellsWritten: 1,
    formulaUtf8Bytes: 0,
    csvPayloadUtf8Bytes: 0,
    commentUtf8Bytes: 0,
    sheetsCreated: 0,
    chartsCreated: 0,
    tablesCreated: 0,
    verificationReads: 1,
  },
  resourceLimits: {
    operationCount: 10,
    cellsRead: 10,
    cellsWritten: 10,
    formulaUtf8Bytes: 1000,
    csvPayloadUtf8Bytes: 1000,
    commentUtf8Bytes: 1000,
    sheetsCreated: 1,
    chartsCreated: 1,
    tablesCreated: 1,
    verificationReads: 10,
  },
};

interface TurnCallCounts {
  model: number;
  firstGroup: number;
  laterGroups: number;
}

function confirmationTurnDependencies(counts: TurnCallCounts = { model: 0, firstGroup: 0, laterGroups: 0 }): WorkbookPlanTurnDependencies {
  const adapter: WorkbookPlanExecutionAdapter & WorkbookPlanVerificationAdapter = {
    async executeFirstGroup(gate) {
      counts.firstGroup += 1;
      return gate.executeAfterFresh("fresh", async () => ({ status: "committed" }));
    },
    async executeGroup() {
      counts.laterGroups += 1;
      return { status: "committed" };
    },
    async reconcileFailedGroup(group) {
      return group.operations.map(() => "unknown_within_batch" as const);
    },
    readVerification() {
      return JSON.stringify({ status: "supported", snapshot: { kind: "range_values", values: [["ok"]] } });
    },
  };
  const response: WorkbookPlanModelResponse = {
    content: "",
    toolCalls: [{
      id: "model-plan",
      type: "function",
      function: { name: "excel_execute_plan", arguments: JSON.stringify(EXECUTE_PLAN) },
    }],
  };
  return {
    identity: IDS,
    trustedOperationBoundary: createTrustedOperationBoundary(IDS),
    compileContext: COMPILE_CONTEXT,
    policyInput: (_plan, confirmation) => ({
      ...CONFIRMATION_REQUIRED_POLICY,
      ...(confirmation === undefined
        ? {}
        : { confirmation: { confirmed: true, eventId: confirmation.eventId } }),
    }),
    store: new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
    model: async () => {
      counts.model += 1;
      return response;
    },
    inspect: async () => "{}",
    resourcePreflightAfterRuntime: async () => ({ ok: true as const }),
    adapter,
    rolloutEnabled: true,
    route: "direct",
    trustedTurnOperationId: IDS.turnId,
  };
}

describe("Task 11 D2 pane-owned pending-plan confirmation", () => {
  it.each(["yes", "proceed", "approved"])(
    "does not treat the initial prompt %s as confirmation authority",
    (prompt) => {
      const request = classifyWorkbookRequest(prompt, {
        eventId: "11111111-1111-4111-8111-111111111111",
        confirmed: true,
      });

      expect(request.confirmation).toBeUndefined();
    },
  );

  it("keeps the D2 pending action content-free and routes pane Continue through one production callback", () => {
    expect(TASKPANE_HTML).toMatch(/id="pending-plan-bar"/);
    expect(TASKPANE_HTML).toMatch(/id="pending-plan-stop"/);
    expect(TASKPANE_HTML).toMatch(/id="pending-plan-continue"/);
    expect(TASKPANE_HTML).not.toMatch(/data-plan|data-digest|workbook-value/i);
    expect(TASKPANE_SOURCE).toContain("continuePendingWorkbookPlan");
    expect(TASKPANE_SOURCE).toMatch(/pendingPlanContinue\.addEventListener\("click", \(\) => continuePendingWorkbookPlan\(\)\)/);
    expect(TASKPANE_SOURCE).toMatch(/pendingPlanStop\.addEventListener\("click", \(\) => stopRunningAgent\(\)\)/);
  });

  it("does not interpolate lower-layer exception text into D2 task-pane failures", () => {
    expect(TASKPANE_SOURCE).not.toMatch(/const message = error instanceof Error \? error\.message/);
    expect(TASKPANE_SOURCE).not.toMatch(/I couldn't complete that request\\n\\n\$\{message\}/);
    // The extracted server-owned route is held to the same content-free rule.
    expect(ROUTE_SOURCE).not.toBe("");
    expect(ROUTE_SOURCE).not.toMatch(/error instanceof Error \? error\.message/);
    expect(ROUTE_SOURCE).not.toMatch(/\$\{message\}/);
  });

  it("routes server-owned pane confirmation through the injected production callback", () => {
    expect(ROUTE_SOURCE).toContain("awaitPaneConfirmation: dependencies.awaitPaneConfirmation");
    expect(ROUTE_SOURCE).toContain("consumePaneConfirmation: (candidate, binding) => dependencies.coordinator.consumeWorkbookPlanConfirmation(candidate, binding)");
    expect(ROUTE_SOURCE).toContain("cancelPaneConfirmation: () => dependencies.coordinator.cancelPendingConfirmation()");
    expect(TASKPANE_SOURCE).toContain("awaitPaneConfirmation: waitForPendingPlanConfirmation");
  });

  it("does not render bridge lower-layer errors into status text", () => {
    expect(TASKPANE_SOURCE).not.toMatch(/Office bridge unavailable: \$\{state\.lastError\}/);
  });

  it("rejects a Continue event minted by a different active pane coordinator", async () => {
    const owner = createCoordinator();
    const foreign = createCoordinator();
    const ownerController = new AbortController();
    const foreignController = new AbortController();
    const ownerHandle = owner.beginTurn(IDS.turnId, ownerController);
    const foreignHandle = foreign.beginTurn(IDS.turnId, foreignController);
    expect((await owner.prepareTurn(ownerHandle)).status).toBe("active");
    expect((await foreign.prepareTurn(foreignHandle)).status).toBe("active");
    const binding: WorkbookPlanConfirmationBinding = {
      planDigest: "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
      sessionId: IDS.sessionId,
      documentId: IDS.documentId,
      turnId: IDS.turnId,
      generation: ownerHandle.generation,
      route: "direct",
      deadline: Date.now() + 1_000,
    };
    const wait = owner.waitForWorkbookPlanConfirmation(binding, ownerHandle.cancellationSignal);
    expect(owner.continuePendingWorkbookPlan()).toBe(true);
    const candidate = await wait;

    expect(foreign.consumeWorkbookPlanConfirmation(candidate, binding)).toBeUndefined();
    expect(owner.consumeWorkbookPlanConfirmation(candidate, binding)).toMatchObject({ confirmed: true, source: "user" });

    owner.finishTurn(ownerHandle);
    foreign.finishTurn(foreignHandle);
  });

  it("consumes only the exact active event once and rejects copied or mismatched candidates", async () => {
    const coordinator = createCoordinator();
    const controller = new AbortController();
    const handle = coordinator.beginTurn(IDS.turnId, controller);
    expect((await coordinator.prepareTurn(handle)).status).toBe("active");
    const binding: WorkbookPlanConfirmationBinding = {
      planDigest: "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
      sessionId: IDS.sessionId,
      documentId: IDS.documentId,
      turnId: IDS.turnId,
      generation: handle.generation,
      route: "direct",
      deadline: Date.now() + 1_000,
    };
    const wait = coordinator.waitForWorkbookPlanConfirmation(binding, handle.cancellationSignal);
    expect(coordinator.continuePendingWorkbookPlan()).toBe(true);
    const candidate = await wait;
    const copied = { ...(candidate as Record<string, unknown>) };
    const wrongBinding = { ...binding, planDigest: "cccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccccc" };

    expect(coordinator.consumeWorkbookPlanConfirmation(copied, binding)).toBeUndefined();
    expect(coordinator.consumeWorkbookPlanConfirmation(candidate, wrongBinding)).toBeUndefined();
    expect(coordinator.consumeWorkbookPlanConfirmation(candidate, binding)).toMatchObject({ confirmed: true, source: "user" });
    expect(coordinator.consumeWorkbookPlanConfirmation(candidate, binding)).toBeUndefined();
    coordinator.finishTurn(handle);
  });

  it("settles a pending confirmation exactly once when the active turn is stopped", async () => {
    const coordinator = createCoordinator();
    const controller = new AbortController();
    const handle = coordinator.beginTurn(IDS.turnId, controller);
    expect((await coordinator.prepareTurn(handle)).status).toBe("active");
    const binding: WorkbookPlanConfirmationBinding = {
      planDigest: "dddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd",
      sessionId: IDS.sessionId,
      documentId: IDS.documentId,
      turnId: IDS.turnId,
      generation: handle.generation,
      route: "direct",
      deadline: Date.now() + 1_000,
    };
    const wait = coordinator.waitForWorkbookPlanConfirmation(binding, handle.cancellationSignal);
    coordinator.stopCurrentTurn();

    await expect(wait).rejects.toMatchObject({ name: "AbortError" });
    expect(coordinator.continuePendingWorkbookPlan()).toBe(false);
  });

  it("rejects an already-expired pending confirmation without creating an event", async () => {
    const coordinator = createCoordinator(() => 2_000);
    const controller = new AbortController();
    const handle = coordinator.beginTurn(IDS.turnId, controller);
    expect((await coordinator.prepareTurn(handle)).status).toBe("active");
    const binding: WorkbookPlanConfirmationBinding = {
      planDigest: "eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee",
      sessionId: IDS.sessionId,
      documentId: IDS.documentId,
      turnId: IDS.turnId,
      generation: handle.generation,
      route: "direct",
      deadline: 2_000,
    };

    await expect(coordinator.waitForWorkbookPlanConfirmation(binding, handle.cancellationSignal)).rejects.toMatchObject({ name: "TimeoutError" });
    expect(coordinator.continuePendingWorkbookPlan()).toBe(false);
    coordinator.finishTurn(handle);
  });

  it("uses the same pending authority for server-owned route bindings", async () => {
    const requested: WorkbookPlanConfirmationBinding[] = [];
    const trusted: TrustedConfirmationEvent = {
      eventId: "66666666-6666-4666-8666-666666666666",
      confirmed: true,
      source: "user",
    };
    const dependencies = {
      ...confirmationTurnDependencies(),
      route: "server-owned",
      trustedTurnOperationId: "77777777-7777-4777-8777-777777777777",
      turnGeneration: 4,
      awaitPaneConfirmation: async (binding: WorkbookPlanConfirmationBinding) => {
        requested.push(binding);
        return Object.freeze({ candidate: true });
      },
      consumePaneConfirmation: () => trusted,
    } as unknown as WorkbookPlanTurnDependencies;

    const result = await createWorkbookPlanTurn(dependencies);

    expect(requested).toHaveLength(1);
    expect(requested[0]).toMatchObject({
      route: "server-owned",
      serverOperationId: "77777777-7777-4777-8777-777777777777",
      generation: 4,
    });
    expect(result.outcome).toBe("committed");
  });

  it("does not hand a confirmation-required plan to the adapter while pane authority is pending", async () => {
    const requested: WorkbookPlanConfirmationBinding[] = [];
    const counts = { model: 0, firstGroup: 0, laterGroups: 0 };
    const candidate = Object.freeze({ candidate: true });
    const trusted: TrustedConfirmationEvent = {
      eventId: "55555555-5555-4555-8555-555555555555",
      confirmed: true,
      source: "user",
    };
    let release!: (value: unknown) => void;
    const authority = new Promise<unknown>((resolve) => {
      release = resolve;
    });
    const dependencies = {
      ...confirmationTurnDependencies(counts),
      awaitPaneConfirmation: async (binding: WorkbookPlanConfirmationBinding) => {
        requested.push(binding);
        return authority;
      },
      consumePaneConfirmation: (value: unknown) => {
        expect(value).toBe(candidate);
        return trusted;
      },
      turnGeneration: 1,
    } as unknown as WorkbookPlanTurnDependencies;

    const resultPromise = createWorkbookPlanTurn(dependencies);
    for (let attempt = 0; attempt < 20 && requested.length === 0; attempt += 1) {
      await new Promise<void>((resolve) => setTimeout(resolve, 0));
    }

    expect(requested).toHaveLength(1);
    expect(counts.model).toBe(1);
    expect(counts.firstGroup).toBe(0);
    release(candidate);
    const result = await resultPromise;

    expect(counts.firstGroup).toBe(1);
    expect(result.outcome).toBe("committed");
  });

  it("keeps a deterministic confirmation-required plan pending until pane authority is supplied", async () => {
    const requested: WorkbookPlanConfirmationBinding[] = [];
    const candidate = Object.freeze({ candidate: true });
    const trusted: TrustedConfirmationEvent = {
      eventId: "44444444-4444-4444-8444-444444444444",
      confirmed: true,
      source: "user",
    };
    const dependencies = {
      ...confirmationTurnDependencies(),
      awaitPaneConfirmation: async (binding: WorkbookPlanConfirmationBinding) => {
        requested.push(binding);
        return candidate;
      },
      consumePaneConfirmation: (value: unknown, binding: WorkbookPlanConfirmationBinding) => {
        expect(value).toBe(candidate);
        expect(binding.planDigest).toMatch(/^[0-9a-f]{64}$/);
        return trusted;
      },
      turnGeneration: 1,
    } as unknown as WorkbookPlanTurnDependencies;

    const result = await createWorkbookPlanTurn(dependencies);

    expect(requested).toHaveLength(1);
    expect(requested[0]).toMatchObject({
      sessionId: IDS.sessionId,
      documentId: IDS.documentId,
      turnId: IDS.turnId,
      generation: 1,
      route: "direct",
    });
    expect(result.outcome).toBe("committed");
  });
});