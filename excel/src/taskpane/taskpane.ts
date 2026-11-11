import DOMPurify from "dompurify";
import { marked } from "marked";
import { HermesAgent, type WorkbookPlanModelContinuation, createTurnCancellation } from "../agent";
import { loadSettings, saveSettings } from "../settings";
import { ActivityItem } from "../types";
import type { WorkbookPlan } from "../generated/workbookPlanSchema";
import { startTrace, finalizeTrace, TurnTrace } from "../turnTracer";
import {
  detectArithmeticRequest,
  formatArithmeticResult,
} from "../deterministicArithmetic";
import { getWorkbookContext, executeExcelTool, parseExcelSelection } from "../excelTools";
import { categorizeError, ToolErrorCategory } from "../toolErrors";
import { classifyPrivacy, PRIVACY_LABELS, PrivacyLevel } from "../privacy";
import { OfficeCompanionBridge, BridgePublicState } from "../officeBridge";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import { createExcelCompanionBridge, executeValidatedExcelBridgeTool, executeValidatedExcelBridgeToolWithAuthority, getOrCreateExcelDocumentId } from "../officeBridgeExcel";
import { createServerPrimitiveGate } from "../primitiveMutationAuthority";
import { createNativeWorkbookPlanAdapter, readNativeWorkbookPlanRuntimeContext } from "../workbookPlanOfficeAdapter";
import {
  createWorkbookPlanTurn,
  ServerOwnedTurnBindingStore,
  type WorkbookPlanConfirmationBinding,
  type WorkbookPlanModelResponse,
} from "../workbookPlanOrchestrator";
import { TaskPaneTurnCoordinator } from "./taskPaneTurnCoordinator";
import { createTrustedOperationBoundary, BrowserIdempotencyStore } from "../workbookPlanExecutor";
import { createReconciliationNotice } from "./reconciliationNotice";

import { classifyWorkbookRequest, type TrustedWorkbookRequest } from "../workbookPlanRequest";
import {
  assertWorkbookPlanRolloutEnabled,
  buildWorkbookPlanPolicyInput,
  preflightWorkbookPlanResources,
  type WorkbookPlanRouteRuntimeContext,
} from "../workbookPlanRoute";
import { isHermesAgentMode, shouldUseLocalArithmeticShortcut } from "../serverOwnedMode";
import {
  createServerOwnedPlanRoute,
  EMPTY_RESOURCE_CONTEXT,
  preflightInitialWorkbookPlanResponse,
  trustedRequestWithConfirmation,
} from "./serverOwnedPlanRoute";
import { browserStorage, colorSchemeMedia, currentOfficeTheme, installThemeControls } from "./omniaTheme";
import { DeskConnector, ApprovalRequiredError, DESK_COMMAND } from "../deskConnect";

declare const __HERMES_OFFICE_ACCEPTANCE__: boolean;

import { bootstrap } from "./bootstrapCoordinator";

let agent!: HermesAgent;
let bridgeLifecycle!: OfficeBridgeLifecycleController;
let bridge: OfficeCompanionBridge | null = null;
let bridgeBase: string | null = null;
let connecting = false;
const connector = new DeskConnector({ client: { host: "Excel" } });

// ---------------------------------------------------------------------------
// UI types and state
// ---------------------------------------------------------------------------

interface UiMessage {
  root: HTMLElement;
  bubble: HTMLElement;
  activity?: HTMLDetailsElement;
  activityList?: HTMLElement;
}

let abortController: AbortController | null = null;
let currentAssistantMessage: UiMessage | null = null;
let assistantBuffer = "";
let settings = loadSettings();
let checkpointFiredThisTurn = false;
// Assigned synchronously at TURN_START and never accepted from model text.
let activeTurnOperationId: string | null = null;
let activeTrustedRequest: TrustedWorkbookRequest | null = null;
let activeTurnStart: number | null = null;
let activeTurnCancellation: { readonly signal: AbortSignal; readonly cleanup: () => void } | null = null;
const directSessionId = crypto.randomUUID();
const serverOwnedTurnBindings = new ServerOwnedTurnBindingStore({
  now: () => Date.now(),
  ttlMs: 90_000,
  maxEntries: 32,
  maxBytes: 64 * 1024,
});
const turnCoordinator = new TaskPaneTurnCoordinator({
  getOrCreateDocumentId: () => getOrCreateExcelDocumentId(),
  getBridgeIdentity: () => ({
    sessionId: bridgeLifecycle?.sessionId ?? null,
    documentId: bridgeLifecycle?.documentId ?? null,
  }),
  reconnectForDocumentChange: (expectedDocumentId) => bridgeLifecycle.reconnectForDocumentChange(expectedDocumentId),
  bindings: serverOwnedTurnBindings,
  paneId: directSessionId,
});
let activeTurnSessionId: string | null = null;
let activeTurnDocumentId: string | null = null;
let activeTurnGeneration: number | null = null;
let lastBridgeSessionId: string | null = null;
let lastBridgeDocumentId: string | null = null;
// Server-owned plan tickets run through the extracted route (R2d-1, M2) with
// the real task-pane dependencies wired here.
const serverPrimitiveGate = createServerPrimitiveGate({
  getBridgeSessionId: () => bridgeLifecycle?.sessionId ?? null,
  getActiveTurn: () => ({ turnOperationId: activeTurnOperationId, sessionId: activeTurnSessionId }),
  resolveAuthority: (input) => turnCoordinator.serverPrimitiveAuthority(input),
  executeWithAuthority: executeValidatedExcelBridgeToolWithAuthority,
  executeReadOnly: executeValidatedExcelBridgeTool,
});
const serverOwnedPlanRoute = createServerOwnedPlanRoute({
  isRolloutEnabled: () => settings.workbookPlanRolloutEnabled === true,
  getBridgeSessionId: () => bridgeLifecycle?.sessionId ?? null,
  getActiveTurn: () => ({
    turnOperationId: activeTurnOperationId,
    trustedRequest: activeTrustedRequest,
    sessionId: activeTurnSessionId,
    turnStart: activeTurnStart,
    hasCancellation: activeTurnCancellation !== null,
  }),
  coordinator: turnCoordinator,
  bindings: serverOwnedTurnBindings,
  readRuntimeContext: (plan) => plan === undefined
    ? readNativeWorkbookPlanRuntimeContext()
    : readNativeWorkbookPlanRuntimeContext(plan),
  createAdapter: () => createNativeWorkbookPlanAdapter(),
  createStore: () => new BrowserIdempotencyStore(),
  inspect: inspectServerWorkbookPlanReads,
  awaitPaneConfirmation: waitForPendingPlanConfirmation,
  executePrimitive: serverPrimitiveGate,
});
// R4i (R3c D4): content-free marker count, read only while this pane is idle. Never awaited by a turn.
const reconciliationNotice = createReconciliationNotice({
  source: { reconciliationRequiredCount: () => new BrowserIdempotencyStore().reconciliationRequiredCount() },
  target: () => document.getElementById("reconciliation-status"),
  isIdle: () => abortController === null && activeTurnOperationId === null,
});

// Checkpoint UI state
let checkpointResolve: ((value: boolean) => void) | null = null;

const compiledAcceptanceEnabled =
  typeof __HERMES_OFFICE_ACCEPTANCE__ !== "undefined" &&
  __HERMES_OFFICE_ACCEPTANCE__ === true;

bootstrap({
  isFullFunctionalityAcceptanceUrl,
  loadAndInstallFullHarness: compiledAcceptanceEnabled
    ? () =>
        import("./fullFunctionalityAcceptanceHarness").then(
          ({ installFullFunctionalityAcceptanceHarness }) =>
            installFullFunctionalityAcceptanceHarness(),
        )
    : () => Promise.resolve(false),
  loadAndInstallFocusedHarness: compiledAcceptanceEnabled
    ? () =>
        import("./officeAcceptanceHarness").then(
          ({ installOfficeAcceptanceHarness }) => {
            installOfficeAcceptanceHarness();
          },
        )
    : () => Promise.resolve(),
  bootTaskpane,
});

function bootTaskpane(): void {
  agent = new HermesAgent();
  bridgeLifecycle = new OfficeBridgeLifecycleController(
    (onState) => {
      bridge = createExcelCompanionBridge(onState, serverOwnedPlanRoute.executeCompanionTool, {
        baseUrl: bridgeBase ?? "", grant: () => connector.grant,
        onApprovalRequired: () => void approvalLost(),
      });
      return bridge;
    },
    updateBridgeState,
  );

  Office.onReady(() => {
    const theme = installThemeControls({ document, storage: browserStorage(), media: colorSchemeMedia(), getOfficeTheme: currentOfficeTheme });
    window.addEventListener("pagehide", (event) => { if (!event.persisted) theme.dispose(); });
    bindControls();
    updatePrivacyBadge();
    setStatus("Ready");
    updateBridgeState({ connected: false, sessionId: null, documentId: null });
    void reconciliationNotice.refresh();
    addAssistantMessage(
      "Hi, I'm Hermes. Ask me to inspect, build, format, calculate, or chart anything in this workbook.",
    );
    void connectToHermes(false);
    bridgeLifecycle.register(
      window,
      Office.addin && typeof Office.addin.onVisibilityModeChanged === "function"
        ? {
            hiddenValue: Office.VisibilityMode.hidden,
            register: (handler: (message: { visibilityMode: unknown }) => void) =>
              Office.addin.onVisibilityModeChanged((message) => {
                if (bridgeBase && connector.grant) handler(message);
              }),
          }
        : undefined,
    );
  });
}

async function runDirectWorkbookPlanTurn(
  initialResponse: WorkbookPlanModelResponse,
  signal: AbortSignal,
  turnStart: number,
  continueModel: WorkbookPlanModelContinuation,
  turnGeneration: number,
  turnController: AbortController,
): Promise<string> {
  assertWorkbookPlanRolloutEnabled(settings.workbookPlanRolloutEnabled);
  const turnId = activeTurnOperationId;
  const trustedRequest = activeTrustedRequest;
  if (turnId === null || trustedRequest === null) throw new Error("Trusted direct turn context is unavailable");
  const turnStartSnapshot = activeTurnStart ?? turnStart;
  const assertLive = (): void => {
    if (!isCurrentTurn(turnId, turnGeneration, turnController, signal)) {
      throw new DOMException("The running Hermes task was stopped.", "AbortError");
    }
  };
  assertLive();
  turnCoordinator.markWorkbookPlanEntered({ generation: turnGeneration, turnId, controller: turnController });
  await preflightInitialWorkbookPlanResponse(initialResponse);
  assertLive();
  const documentId = activeTurnDocumentId ?? await getOrCreateExcelDocumentId();
  assertLive();
  activeTurnDocumentId = documentId;
  let runtime: WorkbookPlanRouteRuntimeContext | undefined;
  const loadRuntime = async (
    plan?: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
    signal?: AbortSignal,
  ): Promise<WorkbookPlanRouteRuntimeContext> => {
    if (signal !== undefined && signal.aborted) {
      throw new DOMException("The running Hermes task was stopped.", "AbortError");
    }
    runtime = plan === undefined
      ? (runtime ?? await readNativeWorkbookPlanRuntimeContext())
      : await readNativeWorkbookPlanRuntimeContext(plan);
    if (signal !== undefined && signal.aborted) {
      throw new DOMException("The running Hermes task was stopped.", "AbortError");
    }
    return runtime;
  };
  const responseModel = async (continuation?: string, modelSignal?: AbortSignal): Promise<WorkbookPlanModelResponse> => {
    if (continuation === undefined) return initialResponse;
    return continueModel(continuation, modelSignal ?? signal);
  };
  const result = await createWorkbookPlanTurn({
    identity: { sessionId: directSessionId, documentId, turnId },
    trustedOperationBoundary: createTrustedOperationBoundary({ sessionId: directSessionId, documentId, turnId }),
    getCompileContextForPlan: async (plan, _signal) => (await loadRuntime(plan, _signal)).compileContext,
    policyInput: (plan, confirmation) => {
      if (runtime === undefined) throw new Error("Trusted workbook runtime context is unavailable");
      return buildWorkbookPlanPolicyInput(plan, runtime, trustedRequestWithConfirmation(trustedRequest, confirmation));
    },
    resourcePreflight: (plan, _signal) => preflightWorkbookPlanResources(plan, EMPTY_RESOURCE_CONTEXT),
    resourcePreflightAfterRuntime: (plan) => runtime === undefined
      ? { ok: false, reason: "live_runtime_unavailable" }
      : preflightWorkbookPlanResources(plan, runtime),
    store: new BrowserIdempotencyStore(),
    model: responseModel,
    inspect: inspectWorkbookPlanReads,
    adapter: createNativeWorkbookPlanAdapter(),
    signal,
    turnStart: turnStartSnapshot,
    rolloutEnabled: settings.workbookPlanRolloutEnabled === true,
    route: "direct",
    trustedTurnOperationId: turnId,
    turnGeneration,
    awaitPaneConfirmation: waitForPendingPlanConfirmation,
    consumePaneConfirmation: (candidate, binding) => turnCoordinator.consumeWorkbookPlanConfirmation(candidate, binding),
    cancelPaneConfirmation: () => turnCoordinator.cancelPendingConfirmation(),
    assertLive,
  });
  return result.finalSummary;
}

async function inspectWorkbookPlanReads(
  reads: readonly { readonly kind: "read_range" | "summarize_range_stats"; readonly sheet: string; readonly address: string }[],
  signal?: AbortSignal,
): Promise<string> {
  const results: unknown[] = [];
  for (const read of reads) {
    if (signal?.aborted === true) throw new DOMException("The running Hermes task was stopped.", "AbortError");
    const raw = await executeValidatedExcelBridgeTool(read.kind, { sheet: read.sheet, address: read.address });
    try { results.push(JSON.parse(raw)); } catch { results.push({ ok: false, error: "Invalid inspection result" }); }
  }
  return JSON.stringify({ ok: true, reads: results });
}

/**
 * R4g-3: the server route's inspection reader. Same reads as the direct route's
 * reader above, then each read in the T8 read shape the route enforces
 * (serverOwnedPlanRoute.ts isBoundedInspectionRead): a tool result object
 * without ok:false becomes {ok:true, data:<that object>}; anything else becomes
 * the fixed {ok:false, data:null}, so no tool error text leaves the pane.
 */
async function inspectServerWorkbookPlanReads(
  reads: readonly { readonly kind: "read_range" | "summarize_range_stats"; readonly sheet: string; readonly address: string }[],
  signal?: AbortSignal,
): Promise<string> {
  const { reads: results } = JSON.parse(await inspectWorkbookPlanReads(reads, signal)) as { readonly reads: readonly unknown[] };
  return JSON.stringify({ ok: true, reads: results.map(toServerInspectionRead) });
}

function toServerInspectionRead(read: unknown): { readonly ok: boolean; readonly data: unknown } {
  return read !== null && typeof read === "object" && !Array.isArray(read) && (read as { readonly ok?: unknown }).ok !== false
    ? { ok: true, data: read }
    : { ok: false, data: null };
}

async function waitForPendingPlanConfirmation(
  binding: WorkbookPlanConfirmationBinding,
  signal?: AbortSignal,
): Promise<unknown> {
  showPendingPlanBar();
  try {
    return await turnCoordinator.waitForWorkbookPlanConfirmation(binding, signal);
  } finally {
    hidePendingPlanBar();
  }
}

function continuePendingWorkbookPlan(): void {
  if (turnCoordinator.continuePendingWorkbookPlan()) {
    hidePendingPlanBar();
    setStatus("Working");
  }
}

function showPendingPlanBar(): void {
  byId<HTMLElement>("pending-plan-bar").classList.remove("hidden");
  setStatus("Awaiting confirmation");
}

function hidePendingPlanBar(): void {
  byId<HTMLElement>("pending-plan-bar").classList.add("hidden");
}


function isFullFunctionalityAcceptanceUrl(): boolean {
  const compiledAcceptanceEnabled =
    typeof __HERMES_OFFICE_ACCEPTANCE__ !== "undefined" &&
    __HERMES_OFFICE_ACCEPTANCE__ === true;
  if (!compiledAcceptanceEnabled || typeof window === "undefined") {
    return false;
  }
  try {
    const url = new URL(window.location.href);
    if (
      url.origin !== "https://localhost:3000" ||
      url.pathname !== "/taskpane.html" ||
      url.username !== "" ||
      url.password !== "" ||
      url.hash !== ""
    ) {
      return false;
    }
    const params = new URLSearchParams(url.search);
    const fullValues = params.getAll("fullFunctionalityAcceptance");
    const hostInfoValues = params.getAll("_host_Info");
    let hasUnknown = false;
    params.forEach((_value, key) => {
      if (key !== "fullFunctionalityAcceptance" && key !== "_host_Info") {
        hasUnknown = true;
      }
    });
    return (
      !hasUnknown &&
      fullValues.length === 1 &&
      fullValues[0] === "1" &&
      (hostInfoValues.length === 0 ||
        (hostInfoValues.length === 1 && hostInfoValues[0] !== ""))
    );
  } catch {
    return false;
  }
}

async function connectToHermes(ask: boolean): Promise<void> {
  if (connecting) return;
  connecting = true;
  try {
    setConnectPanel("searching");
    const base = await connector.discover();
    if (!base) {
      setConnectPanel("no-bridge");
      return;
    }
    bridgeBase = base;
    settings = { ...settings, baseUrl: `${base}/v1` };
    updatePrivacyBadge();
    if (!connector.grant) {
      if (!ask) {
        setConnectPanel("needs-approval");
        return;
      }
      const outcome = await connector.requestApproval(base, (code, command) => setConnectPanel("waiting", code, command));
      if (outcome !== "approved") {
        setConnectPanel(outcome === "denied" ? "denied" : "expired");
        return;
      }
    }
    await bridgeLifecycle.stop();
    await bridgeLifecycle.start();
    setConnectPanel("connected");
  } catch (error) {
    if (error instanceof ApprovalRequiredError) {
      setConnectPanel("needs-approval");
    } else {
      setConnectPanel("error");
    }
  } finally {
    connecting = false;
  }
}

async function approvalLost(): Promise<void> {
  connector.clearGrant();
  await bridgeLifecycle.stop().catch(() => undefined);
  setConnectPanel("needs-approval");
}

type ConnectState = "searching" | "no-bridge" | "needs-approval" | "waiting" | "denied" | "expired" | "connected" | "error";

function setConnectPanel(state: ConnectState, detail = "", command = DESK_COMMAND): void {
  const panel = byId<HTMLElement>("desk-connect");
  const text = byId<HTMLElement>("desk-connect-text");
  const button = byId<HTMLButtonElement>("desk-connect-button");
  const code = byId<HTMLElement>("desk-connect-code");
  panel.hidden = state === "connected";
  code.hidden = state !== "waiting";
  button.disabled = state === "searching" || state === "waiting";
  button.textContent = state === "no-bridge" || state === "error" ? "Try again" : "Connect to Hermes";
  const messages: Record<ConnectState, string> = {
    "searching": "Looking for Hermes on this computer…",
    "no-bridge": `Hermes is not reachable. ${connector.discoveryError || "Start the Hermes gateway with the API server enabled and the Sidecar for Excel plugin installed, then try again."}`,
    "needs-approval": "This sidebar needs your OK in Hermes before it can chat.",
    "waiting": `In any Hermes chat, run ${command} and choose Allow. Hermes shows the same code:`,
    "denied": "Hermes denied this sidebar. Click Connect to Hermes to ask again.",
    "expired": "The request expired. Click Connect to Hermes to ask again.",
    "connected": "",
    "error": "Lost the connection to Hermes on this computer.",
  };
  text.textContent = messages[state];
  code.textContent = state === "waiting" ? detail : "";
  setStatus(state === "connected" ? "Connected to Hermes" : state === "waiting" ? "Waiting for Hermes approval" : "Not connected");
}

function bindControls(): void {
  byId<HTMLButtonElement>("desk-connect-button").addEventListener("click", () => void connectToHermes(true));
  const sendButton = byId<HTMLButtonElement>("send-button");
  const stopButton = byId<HTMLButtonElement>("stop-button");
  const input = byId<HTMLTextAreaElement>("prompt-input");
  const settingsButton = byId<HTMLButtonElement>("settings-button");
  const settingsDialog = byId<HTMLDialogElement>("settings-dialog");
  const saveSettingsButton = byId<HTMLButtonElement>("save-settings-button");
  const copyBridgeButton = byId<HTMLButtonElement>("copy-bridge-session");

  sendButton.addEventListener("click", () => void sendPrompt());
  stopButton.addEventListener("click", () => stopRunningAgent());
  settingsButton.addEventListener("click", () => openSettings());
  saveSettingsButton.addEventListener("click", () => persistSettingsFromDialog());
  copyBridgeButton.addEventListener("click", () => void copyBridgeSessionId());

  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void sendPrompt();
    }
  });

  settingsDialog.addEventListener("cancel", () => fillSettingsDialog());

  // Checkpoint buttons
  const checkpointBar = byId<HTMLElement>("checkpoint-bar");
  const continueBtn = byId<HTMLButtonElement>("checkpoint-continue");
  const stopCheckBtn = byId<HTMLButtonElement>("checkpoint-stop");
  const pendingPlanContinue = byId<HTMLButtonElement>("pending-plan-continue");
  const pendingPlanStop = byId<HTMLButtonElement>("pending-plan-stop");

  continueBtn.addEventListener("click", () => {
    if (checkpointResolve) {
      checkpointResolve(true);
      checkpointResolve = null;
      checkpointBar.classList.add("hidden");
    }
  });

  stopCheckBtn.addEventListener("click", () => {
    if (checkpointResolve) {
      checkpointResolve(false);
      checkpointResolve = null;
      checkpointBar.classList.add("hidden");
    }
  });

  pendingPlanContinue.addEventListener("click", () => continuePendingWorkbookPlan());
  pendingPlanStop.addEventListener("click", () => stopRunningAgent());
}

// ---------------------------------------------------------------------------
// Privacy badge update
// ---------------------------------------------------------------------------

function updatePrivacyBadge(): void {
  const badge = byId<HTMLElement>("privacy-badge");
  const level = classifyPrivacy(settings.baseUrl);
  const info = PRIVACY_LABELS[level];
  badge.textContent = info.text;
  badge.className = "privacy-badge " + info.cssClass;
  badge.title = info.tooltip;
}

async function copyBridgeSessionId(): Promise<void> {
  const sessionId = bridgeLifecycle.sessionId;
  if (!sessionId) return;
  try {
    await navigator.clipboard.writeText(sessionId);
    setStatus("Bridge session ID copied");
  } catch {
    setStatus("Unable to copy bridge session ID");
  }
}

function updateBridgeState(state: BridgePublicState): void {
  if (state.connected) setConnectPanel("connected");
  else if (bridgeBase && !connector.grant) setConnectPanel("needs-approval");
  else if (bridgeBase && state.lastError) setConnectPanel("error");
  const activeTurnInvalidated = turnCoordinator.bridgeStateInvalidatesActiveTurn(
    state,
    activeTurnOperationId === null
      ? null
      : { turnId: activeTurnOperationId, sessionId: activeTurnSessionId, documentId: activeTurnDocumentId },
    { sessionId: lastBridgeSessionId, documentId: lastBridgeDocumentId },
  );
  if (activeTurnInvalidated) stopRunningAgent(true);
  if (lastBridgeSessionId !== null &&
      (lastBridgeSessionId !== state.sessionId || lastBridgeDocumentId !== state.documentId)) {
    serverOwnedTurnBindings.invalidateSession(lastBridgeSessionId);
  }
  if (lastBridgeDocumentId !== null && lastBridgeDocumentId !== state.documentId) {
    serverOwnedTurnBindings.invalidateDocument(lastBridgeDocumentId);
  }
  lastBridgeSessionId = state.sessionId;
  lastBridgeDocumentId = state.documentId;
  const status = byId<HTMLElement>("bridge-status");
  const session = byId<HTMLElement>("bridge-session-id");
  const copy = byId<HTMLButtonElement>("copy-bridge-session");
  if (state.connected && state.sessionId) {
    status.textContent = "Office bridge connected";
    status.className = "bridge-status bridge-connected";
    session.textContent = state.sessionId;
    session.title = state.sessionId;
    copy.disabled = false;
  } else {
    status.textContent = state.lastError
      ? "Office bridge unavailable"
      : "Office bridge connecting";
    status.className = "bridge-status bridge-disconnected";
    session.textContent = "";
    session.removeAttribute("title");
    copy.disabled = true;
  }
}

// ---------------------------------------------------------------------------
// sendPrompt
// ---------------------------------------------------------------------------

async function sendPrompt(): Promise<void> {
  const input = byId<HTMLTextAreaElement>("prompt-input");
  let prompt = input.value.trim();
  if (!prompt || abortController) {
    return;
  }

  const chatHeaders = bridge?.chatHeaders() ?? null;
  if (isHermesAgentMode(settings.model) && (!bridgeBase || !chatHeaders)) {
    setConnectPanel(bridgeBase ? "needs-approval" : "no-bridge");
    return;
  }

  input.value = "";
  addUserMessage(prompt);
  assistantBuffer = "";
  currentAssistantMessage = addAssistantMessage("");
  setRunning(true);

  // Reset per-turn state
  abortController = new AbortController();
  activeTurnOperationId = crypto.randomUUID();
  const turnController = abortController;
  const turnId = activeTurnOperationId;
  const turnStart = Date.now();
  const trustedRequest = classifyWorkbookRequest(prompt);
  const turnCancellation = createTurnCancellation(turnStart, turnController.signal);
  const turnSignal = turnCancellation.signal;
  activeTurnStart = turnStart;
  activeTurnCancellation = turnCancellation;
  activeTurnSessionId = bridgeLifecycle.sessionId;
  activeTurnDocumentId = null;
  activeTrustedRequest = trustedRequest;
  const turnHandle = turnCoordinator.beginTurn(turnId, turnController, {
    trustedRequest,
    turnStart,
    cancellationSignal: turnSignal,
    workbookPlanRolloutEnabled: settings.workbookPlanRolloutEnabled === true,
  });
  const turnGeneration = turnHandle.generation;
  activeTurnGeneration = turnGeneration;
  checkpointFiredThisTurn = false;
  checkpointResolve = null;
  hideCheckpointBar();

  const runTurnContinuation = async (): Promise<void> => {
    // --- Deterministic arithmetic shortcut ---
  const arithmeticReq = detectArithmeticRequest(prompt);
  if (arithmeticReq && shouldUseLocalArithmeticShortcut(settings.model)) {
    try {
      const workbookCtx = await getWorkbookContext();
      const parsed = JSON.parse(workbookCtx) as {
        selection?: string;
        activeSheet?: string;
      };
      const selAddress = parsed.selection || "";
      const { sheet, range } = parseExcelSelection(
        selAddress,
        parsed.activeSheet || "Sheet1",
      );

      const statsJson = await executeExcelTool(
        "summarize_range_stats",
        JSON.stringify({ sheet, address: range }),
      );
      const statsObj = JSON.parse(statsJson);
      const result = formatArithmeticResult(
        arithmeticReq.op,
        statsObj,
        selAddress,
      );
      assistantBuffer = result;
      renderAssistantBuffer();
      setStatus("Ready (local)");
      finishActiveTurn(turnId, turnGeneration, turnController);
      currentAssistantMessage = null;
      return;
    } catch {
      // Fall through to normal agent path
    }
  }

  // Start metadata-only trace
  const trace = startTrace(settings.maxIterations);

  const hermesSessionId = bridgeLifecycle.sessionId;
  const hermesRequest = turnCoordinator.beginHermesRequest(hermesSessionId);
  try {
    const runResult = await agent.run(
      prompt,
      { ...settings, requestHeaders: chatHeaders ?? undefined },
      {
        onAssistantStart: () => {
          assistantBuffer = "";
          renderAssistantBuffer();
        },
        onAssistantDelta: (delta) => {
          assistantBuffer += delta;
          renderAssistantBuffer();
        },
        onAssistantFinal: (content) => {
          assistantBuffer = content;
          renderAssistantBuffer();
        },
        onActivity: (item) => renderActivity(item),
        onStatus: setStatus,
        onCheckpoint: async (deadlineSignal: AbortSignal) => {
          if (turnSignal.aborted) return false;
          if (checkpointFiredThisTurn) return true;
          checkpointFiredThisTurn = true;

          showCheckpointBar();

          return new Promise<boolean>((resolve) => {
            checkpointResolve = resolve;

            const onDeadline = () => {
              if (checkpointResolve === resolve) {
                checkpointResolve = null;
                hideCheckpointBar();
                resolve(false);
              }
            };
            deadlineSignal.addEventListener("abort", onDeadline, {
              once: true,
            });

            const onUserStop = () => {
              if (checkpointResolve === resolve) {
                checkpointResolve = null;
                hideCheckpointBar();
                resolve(false);
              }
            };
            turnSignal.addEventListener("abort", onUserStop, {
              once: true,
            });
          });
        },
        primitiveAuthority: turnCoordinator.primitiveAuthorityFor(turnHandle) ?? undefined,
        onWorkbookPlanTurn: (response, signal, turnStart, continueModel) =>
          runDirectWorkbookPlanTurn(response, signal, turnStart, continueModel, turnGeneration, turnController),
      },
      turnSignal,
      hermesSessionId,
    );

    // Map stopReason (repair 6) to trace outcome
    if (runResult.stopReason === "checkpoint") {
      trace.outcome = "stopped_checkpoint";
      trace.checkpointFired = true;
    } else if (runResult.stopReason === "deadline") {
      trace.outcome = "timeout";
      trace.checkpointFired = true;
    } else {
      trace.outcome = "completed";
    }
    hermesRequest.endClean();
  } catch (error) {
    const category = categorizeError(error);
    trace.errorCategory = category;

    if (error instanceof DOMException && error.name === "AbortError") {
      assistantBuffer = `${assistantBuffer.trim()}\n\nStopped.`;
      renderAssistantBuffer();
      setStatus("Stopped");
      trace.outcome = "stopped_user";
    } else if (
      error instanceof DOMException &&
      error.name === "TimeoutError"
    ) {
      assistantBuffer = `${assistantBuffer.trim()}\n\nTurn exceeded the 90-second time limit and was stopped.`;
      renderAssistantBuffer();
      setStatus("Timeout");
      trace.outcome = "timeout";
    } else {
      assistantBuffer = "I couldn't complete that request.\n\nThe workbook operation was not completed.";
      renderAssistantBuffer();
      setStatus("Error");
      trace.outcome = "error";
      // Content-free error summary: just the class name and category
      trace.errorSummary = error instanceof Error ? error.constructor.name : "Unknown";
    }
  } finally {
    if (hasActiveTurnIdentity(turnId, turnGeneration, turnController)) {
      checkpointResolve = null;
      hideCheckpointBar();
    }

    if (trace) finalizeTrace(trace);
    finishActiveTurn(turnId, turnGeneration, turnController);
  }
  };

  if (settings.workbookPlanRolloutEnabled === true) {
    const gated = await turnCoordinator.continueTurn(turnHandle, async (preparation) => {
      activeTurnDocumentId = preparation.documentId;
      activeTurnSessionId = preparation.sessionId;
      await runTurnContinuation();
    });
    if (gated.status !== "active") {
      finishActiveTurn(turnId, turnGeneration, turnController);
    }
    return;
  }
  await runTurnContinuation();
}

function stopRunningAgent(force = false): void {
  const controller = abortController;
  const turnId = activeTurnOperationId;
  const turnGeneration = activeTurnGeneration;
  if (controller !== null && turnId !== null && turnGeneration !== null &&
      (force || hasActiveTurnIdentity(turnId, turnGeneration, controller))) {
    turnCoordinator.stopCurrentTurn();
    abortController = null;
    activeTurnCancellation?.cleanup();
    activeTurnCancellation = null;
    activeTurnStart = null;
    activeTurnGeneration = null;
    activeTurnOperationId = null;
    activeTrustedRequest = null;
    activeTurnSessionId = null;
    activeTurnDocumentId = null;
    setRunning(false);
  }
  // Also dismiss any pending checkpoint
  if (checkpointResolve) {
    checkpointResolve(false);
    checkpointResolve = null;
    hideCheckpointBar();
  }
}

function hasActiveTurnIdentity(turnId: string, turnGeneration: number, controller: AbortController): boolean {
  return abortController === controller &&
    activeTurnOperationId === turnId &&
    activeTurnGeneration === turnGeneration;
}

function hasCurrentTurn(turnId: string, turnGeneration: number, controller: AbortController): boolean {
  return hasActiveTurnIdentity(turnId, turnGeneration, controller) &&
    turnCoordinator.isCurrent(turnGeneration, turnId, controller);
}

function isCurrentTurn(
  turnId: string,
  turnGeneration: number,
  controller: AbortController,
  signal: AbortSignal,
): boolean {
  return hasCurrentTurn(turnId, turnGeneration, controller) && !signal.aborted;
}

function finishActiveTurn(turnId: string, turnGeneration: number, controller: AbortController): void {
  if (!hasActiveTurnIdentity(turnId, turnGeneration, controller)) return;
  turnCoordinator.finishTurn({
    generation: turnGeneration,
    turnId,
    controller,
    trustedRequest: activeTrustedRequest,
    turnStart: activeTurnStart,
    cancellationSignal: activeTurnCancellation?.signal ?? controller.signal,
  });
  abortController = null;
  activeTurnCancellation?.cleanup();
  activeTurnCancellation = null;
  activeTurnStart = null;
  activeTurnGeneration = null;
  activeTurnOperationId = null;
  activeTrustedRequest = null;
  activeTurnSessionId = null;
  activeTurnDocumentId = null;
  currentAssistantMessage = null;
  setRunning(false);
  void reconciliationNotice.refresh();
}

// ---------------------------------------------------------------------------
// Checkpoint bar UI
// ---------------------------------------------------------------------------

function showCheckpointBar(): void {
  const bar = byId<HTMLElement>("checkpoint-bar");
  bar.classList.remove("hidden");
}

function hideCheckpointBar(): void {
  const bar = byId<HTMLElement>("checkpoint-bar");
  bar.classList.add("hidden");
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

function openSettings(): void {
  fillSettingsDialog();
  byId<HTMLDialogElement>("settings-dialog").showModal();
}

function fillSettingsDialog(): void {
  byId<HTMLElement>("bridge-url").textContent = bridgeBase ?? "not found";
  byId<HTMLInputElement>("model-input").value = settings.model;
  byId<HTMLInputElement>("model-input").readOnly = true;
  byId<HTMLInputElement>("max-iterations-input").value = String(
    settings.maxIterations,
  );
  byId<HTMLInputElement>("workbook-plan-rollout-input").checked =
    settings.workbookPlanRolloutEnabled === true;
}

function persistSettingsFromDialog(): void {
  settings = saveSettings({
    baseUrl: settings.baseUrl,
    model: byId<HTMLInputElement>("model-input").value,
    apiKey: "",
    maxIterations: Number(
      byId<HTMLInputElement>("max-iterations-input").value,
    ),
    workbookPlanRolloutEnabled: byId<HTMLInputElement>(
      "workbook-plan-rollout-input",
    ).checked,
  });
  updatePrivacyBadge();
  setStatus("Settings saved");
}

// ---------------------------------------------------------------------------
// Message rendering
// ---------------------------------------------------------------------------

function addUserMessage(content: string): UiMessage {
  return addMessage("user", content);
}

function addAssistantMessage(content: string): UiMessage {
  return addMessage("assistant", content);
}

function addMessage(
  role: "user" | "assistant",
  content: string,
): UiMessage {
  const messages = byId<HTMLElement>("messages");
  const root = document.createElement("article");
  root.className = `message ${role}`;

  const bubble = document.createElement("div");
  bubble.className = "bubble";
  root.appendChild(bubble);
  messages.appendChild(root);

  const uiMessage = { root, bubble };
  if (role === "assistant") {
    renderMarkdown(bubble, content);
  } else {
    bubble.textContent = content;
  }
  scrollMessagesToBottom();
  return uiMessage;
}

function renderAssistantBuffer(): void {
  if (!currentAssistantMessage) {
    return;
  }
  renderMarkdown(currentAssistantMessage.bubble, assistantBuffer || " ");
  scrollMessagesToBottom();
}

function renderMarkdown(target: HTMLElement, markdown: string): void {
  const html = marked.parse(markdown, {
    async: false,
    breaks: true,
  }) as string;
  target.innerHTML = DOMPurify.sanitize(html);
}

function renderActivity(item: ActivityItem): void {
  if (!currentAssistantMessage) {
    return;
  }

  if (
    !currentAssistantMessage.activity ||
    !currentAssistantMessage.activityList
  ) {
    const activity = document.createElement("details");
    activity.className = "activity";
    activity.open = true;

    const summary = document.createElement("summary");
    summary.textContent = "Activity";
    activity.appendChild(summary);

    const activityList = document.createElement("div");
    activityList.className = "activity-list";
    activity.appendChild(activityList);

    currentAssistantMessage.root.appendChild(activity);
    currentAssistantMessage.activity = activity;
    currentAssistantMessage.activityList = activityList;
  }

  const activityList = currentAssistantMessage.activityList;
  let row = activityList.querySelector<HTMLElement>(
    `[data-activity-id="${item.id}"]`,
  );
  if (!row) {
    row = document.createElement("div");
    row.className = "activity-item";
    row.dataset.activityId = item.id;
    row.innerHTML = "<span></span><span></span>";
    activityList.appendChild(row);
  }

  row.className = `activity-item ${item.status === "error" ? "error" : ""}`;
  row.dataset.status = item.status;
  row.children[0].textContent =
    item.status === "running" ? "\u2026" : item.status === "done" ? "\u2713" : "!";
  row.children[1].textContent = item.label;
  scrollMessagesToBottom();
}

function setRunning(isRunning: boolean): void {
  byId<HTMLButtonElement>("send-button").disabled = isRunning;
  byId<HTMLButtonElement>("stop-button").disabled = !isRunning;
  byId<HTMLTextAreaElement>("prompt-input").disabled = isRunning;
}

function setStatus(status: string): void {
  byId<HTMLElement>("connection-status").textContent = status;
}

function scrollMessagesToBottom(): void {
  const messages = byId<HTMLElement>("messages");
  messages.scrollTop = messages.scrollHeight;
}

function byId<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element #${id}`);
  }
  return element as T;
}
