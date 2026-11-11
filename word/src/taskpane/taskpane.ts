import DOMPurify from "dompurify";
import { marked } from "marked";
import { HermesAgent, AgentCallbacks } from "../agent";
import { loadSettings, saveSettings } from "../settings";
import { ActivityItem } from "../types";
import { classifyPrivacy, PRIVACY_LABELS } from "../privacy";
import { sanitizeErrorBody } from "../toolErrors";
import { BridgePublicState, OfficeCompanionBridge } from "../officeBridge";
import { OfficeBridgeLifecycleController } from "../officeBridgeLifecycle";
import { createWordCompanionBridge } from "../officeBridgeWord";
import { waitForCheckpointDecision } from "./checkpoint";
import { ApprovalRequiredError, DeskConnector, DESK_COMMAND } from "../deskConnect";
declare global {
  interface Window {
    OmniaTheme?: { refresh(): void; bindControls(): void };
  }
}

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
let bridgeBase: string | null = null;
let connecting = false;
let bridge: OfficeCompanionBridge | null = null;
const connector = new DeskConnector({ client: { host: "Word", platform: officePlatform() } });
const bridgeLifecycle = new OfficeBridgeLifecycleController(
  (onState) => {
    bridge = createWordCompanionBridge(onState, {
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

export function formatTaskpaneTimeoutMessage(error: unknown): string {
  const message = sanitizeErrorBody(error instanceof Error ? error.message : String(error), 300);
  return `I couldn't complete that request.\n\n${message}`;
}

export function formatTaskpaneStoppedMessage(currentAssistantBuffer: string): string {
  return `${currentAssistantBuffer.trim()}\n\nStopped.`;
}

Office.onReady(() => {
  window.OmniaTheme?.refresh();
  window.OmniaTheme?.bindControls();
  bindControls();
  updatePrivacyBadge();
  setStatus("Ready");
  updateBridgeState({ connected: false, sessionId: null });
  addAssistantMessage("Hi, I'm Hermes. Ask me to inspect, draft, or format any part of this Word document.");
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
      connector.clearGrant();
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
    "no-bridge": "Hermes is not reachable. Start the Hermes gateway with the API server enabled and the Sidecar for Word plugin installed, then try again.",
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
}

async function sendPrompt(): Promise<void> {
  const input = byId<HTMLTextAreaElement>("prompt-input");
  const submittedDraft = input.value;
  const prompt = input.value.trim();
  if (!prompt || abortController) {
    return;
  }
  const chatHeaders = bridge?.chatHeaders() ?? null;
  if (!bridgeBase || !chatHeaders) {
    setConnectPanel(!connector.grant ? "needs-approval" : bridgeBase ? "error" : "no-bridge");
    return;
  }

  if (input.value === submittedDraft) input.value = "";
  addUserMessage(prompt);
  assistantBuffer = "";
  currentAssistantMessage = addAssistantMessage("");
  setRunning(true);

  abortController = new AbortController();

  const callbacks: AgentCallbacks = {
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
      // Show a nonblocking Continue/Stop prompt in the UI
      setStatus("30s checkpoint — Continue or Stop?");
      const checkpointElement = showCheckpointBanner();
      try {
        return await waitForCheckpointDecision(checkpointElement, deadlineSignal);
      } finally {
        try { checkpointElement.remove(); } catch { /* ignore */ }
      }
    },
  };

  try {
    await agent.run(
      prompt,
      { ...settings, requestHeaders: chatHeaders },
      callbacks,
      abortController.signal,
      bridgeLifecycle.sessionId ?? undefined,
    );
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      assistantBuffer = formatTaskpaneStoppedMessage(assistantBuffer);
      renderAssistantBuffer();
      setStatus("Stopped");
    } else if (error instanceof DOMException && error.name === "TimeoutError") {
      assistantBuffer = formatTaskpaneTimeoutMessage(error);
      renderAssistantBuffer();
      setStatus("Timed out");
    } else {
      const message = sanitizeErrorBody(error instanceof Error ? error.message : String(error), 300);
      assistantBuffer = `I couldn't complete that request.\n\n${message}`;
      renderAssistantBuffer();
      setStatus("Error");
    }
  } finally {
    abortController = null;
    currentAssistantMessage = null;
    setRunning(false);
  }
}

function stopRunningAgent(): void {
  abortController?.abort();
}

function openSettings(): void {
  window.OmniaTheme?.refresh();
  fillSettingsDialog();
  byId<HTMLDialogElement>("settings-dialog").showModal();
}

function fillSettingsDialog(): void {
  byId<HTMLElement>("bridge-url").textContent = bridgeBase ?? "not found";
  byId<HTMLInputElement>("model-input").value = settings.model;
  byId<HTMLInputElement>("model-input").readOnly = true;
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

function updatePrivacyBadge(): void {
  const badge = byId<HTMLElement>("privacy-badge");
  const level = classifyPrivacy(bridgeBase ?? "http://127.0.0.1");
  const label = PRIVACY_LABELS[level];
  badge.textContent = label.text;
  badge.className = `privacy-badge ${label.cssClass}`;
  badge.title = label.tooltip;
}

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
    // Bridge errors can contain URLs or credentials: show fixed copy only.
    status.textContent = state.lastError ? "Office bridge unavailable" : "Office bridge connecting";
    if (!connector.grant) setConnectPanel("needs-approval");
    else if (state.lastError || !connecting) setConnectPanel("error");
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

  row.className = `activity-item ${item.status === "error" ? "error" : item.status === "running" ? "running" : ""}`;
  row.children[0].textContent = item.status === "running" ? "…" : item.status === "done" ? "✓" : "!";
  row.children[1].textContent = item.label;
  scrollMessagesToBottom();
}

function showCheckpointBanner(): HTMLElement {
  const messages = byId<HTMLElement>("messages");
  const banner = document.createElement("div");
  banner.className = "checkpoint-banner";
  banner.setAttribute("role", "dialog");
  banner.setAttribute("aria-live", "assertive");
  banner.setAttribute("aria-label", "Hermes checkpoint");
  banner.innerHTML = `
    <div class="checkpoint-banner-content">
      <span>Hermes has been working for 30 seconds.</span>
      <div class="checkpoint-actions">
        <button class="checkpoint-continue primary-button" type="button">Continue</button>
        <button class="checkpoint-stop secondary-button" type="button">Stop</button>
      </div>
    </div>
  `;
  messages.appendChild(banner);
  (banner.querySelector(".checkpoint-continue") as HTMLButtonElement | null)?.focus();
  scrollMessagesToBottom();
  return banner;
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
