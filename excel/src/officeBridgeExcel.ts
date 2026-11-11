import { executeExcelToolParsed } from "./excelTools";
import {
  OfficeCompanionBridge,
  BridgePublicState,
  type BridgePlanExecutionControl,
  type BridgeTrustedPlanEnvelope,
} from "./officeBridge";
import { validateToolCall } from "./toolValidation";
import {
  type PrimitiveMutationAuthority,
} from "./primitiveMutationAuthority";
import {
  executeAcceptedWorkbookPlan,
  type AcceptedWorkbookPlanTransaction,
  type BrowserIdempotencyStore,
  type WorkbookPlanExecutionAdapter,
  type WorkbookPlanExecutionOptions,
  type WorkbookPlanExecutionReport,
  type TrustedOperationBoundary,
} from "./workbookPlanExecutor";

const DOCUMENT_ID_KEY = "hermes.officeBridge.documentId.v1";

export async function getOrCreateExcelDocumentId(
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
        reject(new Error("Unable to persist the Excel bridge documentId"));
      }
    });
  });
  return documentId;
}

export async function executeValidatedExcelBridgeTool(
  name: string,
  argumentsValue: Record<string, unknown>,
): Promise<string> {
  const validation = validateToolCall(name, JSON.stringify(argumentsValue));
  if (!("valid" in validation) || validation.valid !== true) {
    return JSON.stringify(validation);
  }
  return executeExcelToolParsed(name, validation.parsedArgs);
}

/** R4c: the same validation as executeValidatedExcelBridgeTool, then the turn's primitive authority. */
export async function executeValidatedExcelBridgeToolWithAuthority(
  name: string,
  argumentsValue: Record<string, unknown>,
  authority: PrimitiveMutationAuthority,
): Promise<string> {
  const validation = validateToolCall(name, JSON.stringify(argumentsValue));
  if (!("valid" in validation) || validation.valid !== true) {
    return JSON.stringify(validation);
  }
  return authority.executePrimitive(name, validation.parsedArgs);
}

/** Minimal dependency-injected seam for the plan executor; primitive calls retain their existing path. */
export function executeAcceptedWorkbookPlanThroughExcelBridge(
  transaction: AcceptedWorkbookPlanTransaction,
  trustedOperationBoundary: TrustedOperationBoundary,
  adapter: WorkbookPlanExecutionAdapter,
  store: BrowserIdempotencyStore,
  options?: WorkbookPlanExecutionOptions | AbortSignal,
): Promise<WorkbookPlanExecutionReport> {
  return executeAcceptedWorkbookPlan(transaction, trustedOperationBoundary, adapter, store, options);
}

export {
  getAcceptedWorkbookPlanPostconditionsForVerification,
} from "./workbookPlanExecutor";

/** Optional seams for tests and hosts; production defaults are unchanged. */
export interface ExcelCompanionBridgeOptions {
  readonly fetchImpl?: typeof fetch;
  readonly documentId?: () => Promise<string>;
  readonly paneId?: string;
  readonly baseUrl?: string;
  readonly grant?: () => string | null;
  readonly onApprovalRequired?: () => void;
  readonly setIntervalImpl?: typeof setInterval;
  readonly clearIntervalImpl?: typeof clearInterval;
}

export function createExcelCompanionBridge(
  onState?: (state: BridgePublicState) => void,
  executeTool: (
    name: string,
    argumentsValue: Record<string, unknown>,
    trusted?: BridgeTrustedPlanEnvelope,
    execution?: BridgePlanExecutionControl,
  ) => Promise<string> = executeValidatedExcelBridgeTool,
  options: ExcelCompanionBridgeOptions = {},
): OfficeCompanionBridge {
  return new OfficeCompanionBridge({
    hostPrefix: "excel",
    paneId: options.paneId ?? crypto.randomUUID(),
    documentId: options.documentId ?? (() => getOrCreateExcelDocumentId()),
    executeTool,
    onState,
    grant: options.grant,
    onApprovalRequired: options.onApprovalRequired,
    ...(options.fetchImpl === undefined ? {} : { fetchImpl: options.fetchImpl }),
    ...(options.baseUrl === undefined ? {} : { baseUrl: options.baseUrl }),
    ...(options.setIntervalImpl === undefined ? {} : { setIntervalImpl: options.setIntervalImpl }),
    ...(options.clearIntervalImpl === undefined ? {} : { clearIntervalImpl: options.clearIntervalImpl }),
  });
}

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
