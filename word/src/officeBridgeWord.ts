import { executeWordTool, executeWordToolParsed } from "./wordTools";
import { OfficeCompanionBridge, BridgePublicState } from "./officeBridge";
import { validateToolCall } from "./toolValidation";

const DOCUMENT_ID_KEY = "hermes.officeBridge.documentId.v1";

export async function getOrCreateWordDocumentId(
  randomUuid: () => string = () => crypto.randomUUID(),
): Promise<string> {
  if (typeof Office === "undefined" || !Office.context?.document?.settings) {
    throw new Error("Office document settings are unavailable");
  }
  const settings = Office.context.document.settings;
  const existing = settings.get(DOCUMENT_ID_KEY);
  if (typeof existing === "string" && isUuid(existing)) {
    return existing;
  }

  const documentId = randomUuid();
  if (!isUuid(documentId)) {
    throw new Error("Generated documentId must be a UUID");
  }
  settings.set(DOCUMENT_ID_KEY, documentId);
  await new Promise<void>((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error("Unable to persist the Word bridge documentId"));
      }
    });
  });
  return documentId;
}

export async function executeValidatedWordBridgeTool(
  name: string,
  argumentsValue: Record<string, unknown>,
): Promise<string> {
  const validation = validateToolCall(name, JSON.stringify(argumentsValue));
  if (!("valid" in validation) || validation.valid !== true) {
    if (["find_replace", "apply_edit_plan", "apply_style", "insert_text"].includes(name)) {
      return executeWordTool(name, JSON.stringify(argumentsValue));
    }
    return JSON.stringify(validation);
  }
  return executeWordToolParsed(name, validation.parsedArgs);
}

export interface DeskBridgeOptions {
  /** The Sidecar bridge found on 127.0.0.1 (DeskConnector.discover). */
  baseUrl: string;
  grant: () => string | null;
  onApprovalRequired?: () => void;
}

export function createWordCompanionBridge(
  onState?: (state: BridgePublicState) => void,
  desk?: DeskBridgeOptions,
): OfficeCompanionBridge {
  return new OfficeCompanionBridge({
    hostPrefix: "word",
    paneId: crypto.randomUUID(),
    documentId: () => getOrCreateWordDocumentId(),
    executeTool: executeValidatedWordBridgeTool,
    onState,
    baseUrl: desk?.baseUrl,
    grant: desk?.grant,
    onApprovalRequired: desk?.onApprovalRequired,
  });
}

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
