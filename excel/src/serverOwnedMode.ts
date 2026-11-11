import { ChatMessage } from "./types";
import type { HermesSettings } from "./settings";

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export function isValidBridgeSessionId(value: unknown): value is string {
  return typeof value === "string" && UUID_PATTERN.test(value);
}

/**
 * Returns true when `model` is exactly "hermes-agent".
 * In that mode the Hermes gateway runs its own server-side native plugin loop.
 * Caller-sent tools are silently ignored by the API; we must not send them.
 */
export function isHermesAgentMode(model: string): boolean {
  return model.trim() === "hermes-agent";
}

/** Keep every task-pane tool path disabled while Hermes owns the agent loop. */
export function shouldUseLocalArithmeticShortcut(model: string): boolean {
  return !isHermesAgentMode(model);
}

/**
 * Build a chat-completions body for server-owned (native plugin) mode.
 * Sends workbook context, no tools/tool_choice.
 * Fails closed if there is no bridge session — hermes-agent requires one.
 */
export function buildServerOwnedRequest(
  settings: HermesSettings,
  messages: ChatMessage[],
  bridgeSessionId: string | null,
): { model: string; messages: ChatMessage[]; stream: boolean; temperature: number } {
  if (!isValidBridgeSessionId(bridgeSessionId)) {
    throw new Error(
      "Hermes-agent mode requires an active Office bridge session with a valid UUID. " +
        "Open the Sidecar for Excel task pane and try again.",
    );
  }

  return {
    model: settings.model,
    messages,
    stream: true,
    temperature: 0.2,
  };
}
