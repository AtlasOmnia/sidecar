import type { HermesSettings } from "./settings";
import type { ChatMessage } from "./types";

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export const WORD_EDIT_SAFETY_GUIDANCE = `Await each mutation before reading back actual text, paragraph boundaries, and styles with get_document_structure and read_document_text. A write receipt or matching text alone does not prove a style. Use insert_text paragraphs mode for headings/body with explicit native paragraph boundaries and built-in styles; do not assemble them from incomplete OOXML fragments. For style-only repair, use apply_style with the exact complete targetText and verify observed styles, independent of current selection.
Direct searches are bounded to 255 UTF-16 units in both source and encoded literal query; replacements remain 5000. Never truncate or split sources to evade the bound. On a refused or failed search, stop destructive edits, inspect and report; earlier successful edits remain. Do not delete/reinsert, flatten the document, or blindly replay edits as recovery. If mutationState is unknown, stop and ask the user to inspect/use native Undo; do not retry or compensate automatically. These instructions are behavioral guidance, not a hard execution fence.`;

const SERVER_OWNED_SYSTEM_PROMPT = `You are Hermes, an agentic AI assistant running inside the user's active Microsoft Word document.

The Hermes server owns this entire agent loop. Use the native sidecar_word_* tools available to you. Those tools are automatically bound to the active Word bridge session. Execute document actions through those native tools and use their results before answering.

${WORD_EDIT_SAFETY_GUIDANCE}

Do not emit tool calls as text, JSON, or XML. Do not ask the task pane to execute a tool. When the requested document action has completed, return only a concise final summary.`;

export function isHermesAgentMode(model: string): boolean {
  return model.trim() === "hermes-agent";
}

export function isValidBridgeSessionId(value: unknown): value is string {
  return typeof value === "string" && UUID_PATTERN.test(value);
}

export function getServerOwnedSystemPrompt(): string {
  return SERVER_OWNED_SYSTEM_PROMPT;
}

export function buildServerOwnedRequest(
  settings: HermesSettings,
  messages: ChatMessage[],
  bridgeSessionId: string | null | undefined,
): { model: string; messages: ChatMessage[]; stream: boolean; temperature: number } {
  if (!isValidBridgeSessionId(bridgeSessionId)) {
    throw new Error(
      "Hermes-agent mode requires an active Word bridge session with a valid UUID. Open the Sidecar for Word task pane and try again.",
    );
  }

  return {
    model: settings.model,
    messages,
    stream: true,
    temperature: 0.2,
  };
}
