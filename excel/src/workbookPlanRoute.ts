import type { WorkbookOperation, WorkbookPlan } from "./generated/workbookPlanSchema";
import {
  WORKBOOK_PLAN_CAPABILITY_NAMES,
  type WorkbookPlanOperationEffect,
  type WorkbookPlanOperationSummary,
  type WorkbookPlanPolicyInput,
  type WorkbookPlanSelectionRelationship,
} from "./workbookPlanPolicy";
import type { VerifiedWorkbookCompileContext } from "./workbookPlanCompiler";
import {
  accountWorkbookPlanResources,
  type WorkbookPlanResourceContext,
} from "./workbookPlanResources";
import type { TrustedWorkbookRequest } from "./workbookPlanRequest";
import { workbookPlanContract } from "./workbookPlanContract";

export interface WorkbookPlanRouteSelection {
  readonly sheetName: string;
  readonly address: string;
}

export interface WorkbookPlanRouteRuntimeContext extends WorkbookPlanResourceContext {
  readonly activeSheetName: string;
  readonly selection?: WorkbookPlanRouteSelection;
  readonly compileContext: VerifiedWorkbookCompileContext;
}

const ADDRESS_RE = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?$/i;

interface Bounds {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

function columnNumber(value: string): number {
  let result = 0;
  for (const character of value.toUpperCase()) result = result * 26 + character.charCodeAt(0) - 64;
  return result;
}

function bounds(address: string): Bounds | undefined {
  const match = ADDRESS_RE.exec(address.trim());
  if (match === null) return undefined;
  const firstRow = Number(match[2]);
  const lastRow = Number(match[4] ?? match[2]);
  const firstColumn = columnNumber(match[1]);
  const lastColumn = columnNumber(match[3] ?? match[1]);
  return {
    top: Math.min(firstRow, lastRow),
    left: Math.min(firstColumn, lastColumn),
    bottom: Math.max(firstRow, lastRow),
    right: Math.max(firstColumn, lastColumn),
  };
}

function operationSheets(
  operation: WorkbookOperation,
  compileContext: VerifiedWorkbookCompileContext,
): readonly string[] {
  switch (operation.kind) {
    case "create_sheet": return [operation.name];
    case "rename_sheet": return [operation.oldName, operation.newName];
    case "copy_sheet": return [operation.sourceName, operation.newName];
    case "sort_table":
    case "filter_table":
    case "reset_table_filter": {
      const sheetName = compileContext.tablesByName.get(operation.tableName)?.sheetName;
      return sheetName === undefined ? [] : [sheetName];
    }
    case "calculate": return [];
    default: return [operation.sheet];
  }
}

function operationAddress(operation: WorkbookOperation): string | undefined {
  switch (operation.kind) {
    case "fill_formula": return operation.targetAddress;
    case "create_chart": return operation.dataRange;
    case "calculate":
    case "create_sheet":
    case "rename_sheet":
    case "copy_sheet":
    case "sort_table":
    case "filter_table":
    case "reset_table_filter":
    case "freeze_panes":
      return undefined;
    default: return operation.address;
  }
}

function selectionRelationship(
  operation: WorkbookOperation,
  sheetNames: readonly string[],
  runtime: WorkbookPlanRouteRuntimeContext,
): WorkbookPlanSelectionRelationship {
  const address = operationAddress(operation);
  // A table/sheet/pane mutation without a proven bounded address is not
  // selection-independent. Unknown scope requires confirmation for a
  // selection-scoped request; only non-mutating calculation is inapplicable.
  if (address === undefined) return operation.kind === "calculate" ? "not_applicable" : "unknown";
  const selection = runtime.selection;
  if (selection === undefined || sheetNames.length !== 1) return "unknown";
  if (sheetNames[0] !== selection.sheetName) return "outside";
  const target = bounds(address);
  const selected = bounds(selection.address);
  if (target === undefined || selected === undefined) return "unknown";
  return target.top >= selected.top &&
    target.left >= selected.left &&
    target.bottom <= selected.bottom &&
    target.right <= selected.right
    ? "inside"
    : "outside";
}

export function buildWorkbookPlanPolicyInput(
  plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
  runtime: WorkbookPlanRouteRuntimeContext,
  trustedRequest: TrustedWorkbookRequest,
): WorkbookPlanPolicyInput {
  const accounted = accountWorkbookPlanResources(plan, runtime);
  if (accounted.ok !== true) throw new Error(`Workbook plan resource preflight failed: ${accounted.reason}`);

  const operationSummaries: WorkbookPlanOperationSummary[] = plan.operations.map((operation, operationIndex) => {
    const definition = workbookPlanContract.operationDefinitions.find((candidate) => candidate.kind === operation.kind);
    if (definition === undefined) throw new Error("Workbook plan operation metadata is unavailable");
    const sheetNames = operationSheets(operation, runtime.compileContext);
    return {
      operationIndex,
      effect: definition.effect as WorkbookPlanOperationEffect,
      target: {
        operationIndex,
        sheetNames,
        selectionRelationship: selectionRelationship(operation, sheetNames, runtime),
      },
    };
  });

  return {
    requestClassification: trustedRequest.classification,
    activeSheetName: runtime.activeSheetName,
    selectionScopedRequest: trustedRequest.selectionScopedRequest,
    operationSummaries,
    capabilitySummaries: [...WORKBOOK_PLAN_CAPABILITY_NAMES],
    resourceUsage: accounted.usage,
    resourceLimits: accounted.limits,
    ...(trustedRequest.confirmation
      ? {
          confirmation: {
            confirmed: true,
            eventId: trustedRequest.confirmation.eventId,
          },
        }
      : {}),
  };
}

export function assertWorkbookPlanRolloutEnabled(enabled: unknown): void {
  if (enabled !== true) throw new Error("Workbook plan execution is disabled by the rollout setting");
}

export function preflightWorkbookPlanResources(
  plan: WorkbookPlan,
  runtime: WorkbookPlanResourceContext,
): { readonly ok: true } | { readonly ok: false; readonly reason: string } {
  const result = accountWorkbookPlanResources(plan, runtime);
  // The route's initial context is intentionally empty. Unknown table bounds
  // are deferred to the plan-aware live context; all other conservative
  // failures remain early rejections. An empty context never becomes the
  // final execution authority because the orchestrator re-runs this check
  // after runtime capture.
  if (
    result.ok !== true &&
    result.reason === "unknown_table_bounds" &&
    runtime.compileContext.sheetNames.size === 0 &&
    runtime.compileContext.tablesByName.size === 0 &&
    runtime.compileContext.supportedPrimitiveNames.size === 0 &&
    runtime.compileContext.supportedCapabilities.size === 0 &&
    runtime.compileContext.chartSourceResolutionsBySheet.size === 0
  ) {
    return { ok: true };
  }
  return result.ok === true ? { ok: true } : { ok: false, reason: result.reason };
}
