import { executePowerPointToolParsed } from "./powerPointTools";
import { OfficeCompanionBridge, BridgePublicState } from "./officeBridge";
import { validateToolCall } from "./toolValidation";

const DOCUMENT_ID_KEY = "hermes.officeBridge.documentId.v1";

export async function getOrCreatePowerPointDocumentId(
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
        reject(new Error("Unable to persist the PowerPoint bridge documentId"));
      }
    });
  });
  return documentId;
}

export async function executeValidatedPowerPointBridgeTool(
  name: string,
  argumentsValue: Record<string, unknown>,
): Promise<string> {
  const validation = validateToolCall(name, JSON.stringify(argumentsValue));
  if (!("valid" in validation) || validation.valid !== true) {
    return JSON.stringify(validation);
  }
  return executePowerPointToolParsed(name, validation.parsedArgs);
}

export interface DeskBridgeOptions {
  baseUrl: string;
  grant: () => string | null;
  onApprovalRequired?: () => void;
}

export function createPowerPointCompanionBridge(
  onState?: (state: BridgePublicState) => void,
  desk?: DeskBridgeOptions,
): OfficeCompanionBridge {
  return new OfficeCompanionBridge({
    hostPrefix: "powerpoint",
    paneId: crypto.randomUUID(),
    documentId: () => getOrCreatePowerPointDocumentId(),
    executeTool: executeValidatedPowerPointBridgeTool,
    onState,
    baseUrl: desk?.baseUrl,
    grant: desk?.grant,
    onApprovalRequired: desk?.onApprovalRequired,
  });
}

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
