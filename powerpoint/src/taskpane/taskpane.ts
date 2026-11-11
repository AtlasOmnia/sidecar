import DOMPurify from "dompurify";
import { marked } from "marked";
import { HermesAgent } from "../agent";
import { loadSettings, saveSettings } from "../settings";
import { ActivityItem } from "../types";
import { classifyPrivacy, PRIVACY_LABELS } from "../privacy";
import { startTrace, finalizeTrace, recordModelCall, recordToolResult, recordCheckpoint } from "../turnTracer";
import { CheckpointController } from "../checkpoint";
import { asToolError } from "../toolErrors";
import { BridgePublicState, OfficeCompanionBridge } from "../officeBridge";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import { createPowerPointCompanionBridge } from "../officeBridgePowerPoint";
import { ApprovalRequiredError, DeskConnector, DESK_COMMAND } from "../deskConnect";
// Styles and the theme bootstrap load in <head> before the task-pane bundle.

interface UiMessage {
  root: HTMLElement;
  bubble: HTMLElement;
  activity?: HTMLDetailsElement;
  activityList?: HTMLElement;
}

const agent = new HermesAgent();
let abortController: AbortController | null = null;
let currentAssistantMessage: UiMessage | null = null;
let assistantBuffer = "";
let settings = loadSettings();
const checkpoint = new CheckpointController();
let bridgeBase: string | null = null;
let connecting = false;
let bridge: OfficeCompanionBridge | null = null;
const connector = new DeskConnector({ client: { host: "PowerPoint", platform: officePlatform() } });
const bridgeLifecycle = new OfficeBridgeLifecycleController(
  (onState) => {
    bridge = createPowerPointCompanionBridge(onState, {
      baseUrl: bridgeBase ?? "",
      grant: () => connector.grant,
      onApprovalRequired: () => void approvalLost(),
    });
    return bridge;
  },
  updateBridgeState,
);

function officePlatform(): string {
  try {
    return String(Office?.context?.platform ?? "");
  } catch {
    return "";
  }
}

Office.onReady(() => {
  window.OmniaTheme?.observeOffice();
  bindControls();
  updatePrivacyBadge();
  setStatus("Ready");
  updateBridgeState({ connected: false, sessionId: null });
  addAssistantMessage("Hi, I'm Hermes. Ask me to inspect, draft, or edit the active presentation.");
  bridgeLifecycle.register(
    window,
    Office.addin && typeof Office.addin.onVisibilityModeChanged === "function"
      ? {
          hiddenValue: Office.VisibilityMode.hidden,
          register: (handler) => Office.addin.onVisibilityModeChanged((message) => {
            if (bridgeBase) handler(message);
          }),
        }
      : undefined,
  );
  void connectToHermes(false);
});

/**
 * Find the Sidecar bridge, then start a session with the remembered approval. With no approval yet,
 * show "Connect to Hermes"; when `ask` is true (the user clicked it), file the request and wait for the
 * user's Allow/Deny in Hermes.
 */
async function connectToHermes(ask: boolean): Promise<void> {
  if (connecting) return;
  connecting = true;
  try {
    setConnectPanel("searching");
    const base = await connector.discover();
    if (!base) {
      setConnectPanel(connector.grant ? "error" : "no-bridge");
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
    "no-bridge": "Hermes is not reachable. Start the Hermes gateway with the API server enabled and the Sidecar for PowerPoint plugin installed, trust its local CA, then try again.",
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
  const sendButton = byId<HTMLButtonElement>("send-button");
  const stopButton = byId<HTMLButtonElement>("stop-button");
  const input = byId<HTMLTextAreaElement>("prompt-input");
  const settingsButton = byId<HTMLButtonElement>("settings-button");
  const settingsDialog = byId<HTMLDialogElement>("settings-dialog");
  const saveSettingsButton = byId<HTMLButtonElement>("save-settings-button");
  const copyBridgeButton = byId<HTMLButtonElement>("copy-bridge-session");

  sendButton.addEventListener("click", () => void sendPrompt());
  stopButton.addEventListener("click", stopRunningAgent);
  settingsButton.addEventListener("click", () => openSettings());
  saveSettingsButton.addEventListener("click", () => persistSettingsFromDialog());
  copyBridgeButton.addEventListener("click", () => void copyBridgeSessionId());
  byId<HTMLButtonElement>("desk-connect-button").addEventListener("click", () => void connectToHermes(true));

  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      void sendPrompt();
    }
  });

  settingsDialog.addEventListener("cancel", () => fillSettingsDialog());
  byId<HTMLButtonElement>("checkpoint-continue").addEventListener("click", () => { checkpoint.choose("continue"); byId<HTMLElement>("checkpoint-bar").classList.add("hidden"); });
  byId<HTMLButtonElement>("checkpoint-stop").addEventListener("click", () => { checkpoint.choose("stop"); byId<HTMLElement>("checkpoint-bar").classList.add("hidden"); });
}

// ---------------------------------------------------------------------------
// Companion bridge lifecycle
// ---------------------------------------------------------------------------

function updateBridgeState(state: BridgePublicState): void {
  const status = byId<HTMLElement>("bridge-status");
  const session = byId<HTMLElement>("bridge-session-id");
  const copy = byId<HTMLButtonElement>("copy-bridge-session");
  if (state.connected && state.sessionId) {
    setConnectPanel("connected");
    status.textContent = "Office bridge connected";
    status.className = "bridge-status bridge-connected";
    session.textContent = state.sessionId;
    session.title = state.sessionId;
    copy.disabled = false;
  } else {
    status.textContent = state.lastError
      ? "Office bridge unavailable"
      : "Office bridge connecting";
    if (state.lastError) setConnectPanel(connector.grant ? "error" : "needs-approval");
    status.className = "bridge-status bridge-disconnected";
    session.textContent = "";
    session.removeAttribute("title");
    copy.disabled = true;
  }
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

async function sendPrompt(): Promise<void> {
  const input = byId<HTMLTextAreaElement>("prompt-input");
  const prompt = input.value.trim();
  if (!prompt || abortController) {
    return;
  }

  const chatHeaders = bridge?.chatHeaders() ?? null;
  if (!bridgeBase || !chatHeaders) {
    setConnectPanel(!connector.grant ? "needs-approval" : "error");
    return;
  }

  const bridgeSessionId = bridgeLifecycle.sessionId;
  input.value = "";
  addUserMessage(prompt);
  assistantBuffer = "";
  currentAssistantMessage = addAssistantMessage("");
  setRunning(true);

  abortController = new AbortController();
  const trace = startTrace(settings.maxIterations);

  try {
    const runResult = await agent.run(prompt, { ...settings, requestHeaders: chatHeaders }, {
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
      onModelCall: (event) => recordModelCall(trace, event),
      onToolResult: (event) => recordToolResult(trace, event.name, event.durationMs, event.ok, event.category),
      onCheckpointResult: (decision) => recordCheckpoint(trace, decision),
      onCheckpoint: async (deadlineSignal) => {
        const bar = byId<HTMLElement>("checkpoint-bar");
        bar.classList.remove("hidden");
        try {
          return await checkpoint.begin(deadlineSignal);
        } finally {
          bar.classList.add("hidden");
        }
      }
    }, abortController.signal, { bridgeSessionId });
    trace.outcome = runResult.stopReason || "completed";
  } catch (error) {
    trace.outcome = error instanceof DOMException ? error.name : "error";
    if (error instanceof DOMException && error.name === "AbortError") {
      assistantBuffer = `${assistantBuffer.trim()}\n\nStopped.`;
      renderAssistantBuffer();
      setStatus("Stopped");
    } else {
      const message = formatAgentError(error);
      assistantBuffer = `I couldn't complete that request.\n\n${message}`;
      renderAssistantBuffer();
      setStatus("Error");
    }
  } finally {
    checkpoint.cancel();
    byId<HTMLElement>("checkpoint-bar").classList.add("hidden");
    finalizeTrace(trace);
    abortController = null;
    currentAssistantMessage = null;
    setRunning(false);
  }
}

export function formatAgentError(error: unknown): string {
  const toolError = error && typeof error === "object" && "toolError" in error
    ? (error as { toolError?: unknown }).toolError
    : undefined;
  return asToolError(toolError ?? error).error;
}

function stopRunningAgent(): void {
  abortController?.abort();
}

function openSettings(): void {
  fillSettingsDialog();
  byId<HTMLDialogElement>("settings-dialog").showModal();
}

function fillSettingsDialog(): void {
  byId<HTMLElement>("bridge-url").textContent = bridgeBase ?? "not found";
  byId<HTMLInputElement>("model-input").value = settings.model;
  byId<HTMLInputElement>("max-iterations-input").value = String(settings.maxIterations);
}

function persistSettingsFromDialog(): void {
  settings = saveSettings({
    baseUrl: settings.baseUrl,
    model: byId<HTMLInputElement>("model-input").value,
    maxIterations: Number(byId<HTMLInputElement>("max-iterations-input").value)
  });
  updatePrivacyBadge();
  setStatus("Settings saved");
}

function addUserMessage(content: string): UiMessage {
  return addMessage("user", content);
}

function addAssistantMessage(content: string): UiMessage {
  return addMessage("assistant", content);
}

function addMessage(role: "user" | "assistant", content: string): UiMessage {
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
  const html = marked.parse(markdown, { async: false, breaks: true }) as string;
  target.innerHTML = DOMPurify.sanitize(html);
}

function renderActivity(item: ActivityItem): void {
  if (!currentAssistantMessage) {
    return;
  }

  if (!currentAssistantMessage.activity || !currentAssistantMessage.activityList) {
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
  let row = activityList.querySelector<HTMLElement>(`[data-activity-id="${item.id}"]`);
  if (!row) {
    row = document.createElement("div");
    row.className = "activity-item";
    row.dataset.activityId = item.id;
    row.innerHTML = "<span></span><span></span>";
    activityList.appendChild(row);
  }

  row.className = `activity-item ${item.status === "error" ? "error" : ""}`;
  row.dataset.status = item.status;
  row.children[0].textContent = item.status === "running" ? "\u2026" : item.status === "done" ? "\u2713" : "!";
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

function updatePrivacyBadge(): void {
  const badge = document.getElementById("privacy-badge");
  if (!badge) return;
  const privacy = classifyPrivacy(bridgeBase ?? "http://127.0.0.1");
  badge.textContent = PRIVACY_LABELS[privacy].label;
  badge.title = PRIVACY_LABELS[privacy].tooltip;
  badge.dataset.privacy = privacy;
}
