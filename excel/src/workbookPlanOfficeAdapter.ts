import type { VerificationRule, ScalarValue, ValueMatrix, WorkbookPlan } from "./generated/workbookPlanSchema";
import type {
  WorkbookPlanEffectGroup,
  WorkbookPlanEffectOperation,
  WorkbookPlanFirstGroupContext,
  WorkbookPlanGroupExecutionResult,
  WorkbookPlanExecutionAdapter,
} from "./workbookPlanExecutor";
import type {
  VerifiedWorkbookCompileContext,
  WorkbookPlanPreconditionManifest,
  WorkbookPlanPreconditionRange,
  CompiledWorkbookMandatoryPostcondition,
} from "./workbookPlanCompiler";
import { workbookPlanContract } from "./workbookPlanContract";
import {
  prepareWorkbookPlanVerificationRules,
  verifyWorkbookPlan,
  type WorkbookPlanVerificationReport,
} from "./workbookPlanVerification";

type OfficeRange = {
  values?: unknown;
  formulas?: unknown;
  numberFormat?: unknown;
  address?: string;
  rowCount?: number;
  columnCount?: number;
  format?: Record<string, unknown> & { load?: (fields: string | readonly string[]) => void; font?: Record<string, unknown> & { load?: (fields: string | readonly string[]) => void }; fill?: Record<string, unknown> & { load?: (fields: string | readonly string[]) => void }; borders?: Record<string, unknown> & { load?: (fields: string | readonly string[]) => void; items?: readonly Record<string, unknown>[] } };
  load?: (fields: string | readonly string[]) => void;
  getCell?: (row: number, column: number) => OfficeRange;
  getResizedRange?: (rows: number, columns: number) => OfficeRange;
  autoFill?: (destination: OfficeRange) => void;
  dataValidation?: Record<string, unknown>;
  conditionalFormats?: { add: (type: string) => Record<string, unknown>; load?: (fields: string | readonly string[]) => void; items?: readonly Record<string, unknown>[] };
};
type OfficeRangeAreas = {
  address?: string;
  areaCount?: number;
  load?: (fields: string | readonly string[]) => void;
};
type OfficeWorksheet = {
  getRange: (address: string) => OfficeRange;
  getRanges?: (address: string) => OfficeRangeAreas;
  load?: (fields: string | readonly string[]) => void;
  getUsedRangeOrNullObject?: () => OfficeRange & { isNullObject?: boolean };
  name?: string;
  charts?: {
    add: (type: string, source: OfficeRange, seriesBy: string) => OfficeChart;
    load?: (fields: string | readonly string[]) => void;
    items?: OfficeChart[];
    getItemOrNullObject?: (title: string) => OfficeChart & { isNullObject?: boolean };
  };
  tables?: { add: (range: OfficeRange, hasHeaders: boolean) => OfficeTable };
  comments?: OfficeComments;
  freezePanes?: { freezeRows: (count: number) => void; freezeColumns: (count: number) => void; getLocationOrNullObject?: () => OfficeRange & { isNullObject?: boolean } };
  copy?: (position: string, relativeTo?: OfficeWorksheet) => OfficeWorksheet;
};
type OfficeChart = { name?: string; chartType?: string; title?: { text?: string }; left?: number; top?: number; width?: number; height?: number; load?: (fields: string | readonly string[]) => void; isNullObject?: boolean };
type OfficeComment = { content?: string; load?: (fields: string | readonly string[]) => void; getLocation: () => OfficeRange };
/** Narrow OfficeExtension.ClientResult shape; value is readable only after context.sync(). */
type ClientResult<T> = { readonly value: T };
/** Excel.CommentCollection subset used by this adapter. getCount bounds all item enumeration. */
type OfficeComments = {
  items?: OfficeComment[];
  getCount: () => ClientResult<number>;
  load: (fields: string | readonly string[]) => void;
  add: (range: OfficeRange, text: string) => OfficeComment;
};
type OfficeCommentCandidate = { readonly comment: OfficeComment; readonly location: OfficeRange };
type OfficeCommentRequest = { readonly sheet: string; readonly comments: OfficeComments };
type OfficeCommentScan = {
  readonly comments: OfficeComments;
  readonly count: number;
  readonly items: readonly OfficeComment[];
  readonly candidates: readonly OfficeCommentCandidate[];
  readonly index: ReadonlyMap<string, OfficeComment>;
};
type OfficeCommentScans = ReadonlyMap<string, OfficeCommentScan>;

const MAX_SORT_BODY_ROWS = 10_000;
const MAX_SORT_BODY_COLUMNS = 1_000;
const MAX_SORT_BODY_CELLS = 100_000;
const MAX_FILTER_VALUE_LENGTH = 512;
type OfficeContext = {
  workbook: {
    getSelectedRange?: () => OfficeRange;
    worksheets: {
      getItem: (name: string) => OfficeWorksheet;
      add?: (name: string) => OfficeWorksheet;
      getItemOrNullObject?: (name: string) => OfficeWorksheet & { isNullObject?: boolean };
      getActiveWorksheet?: () => OfficeWorksheet;
      getSelectedRange?: () => OfficeRange;
      load?: (fields: string | readonly string[]) => void;
      items?: OfficeWorksheet[];
    };
    tables?: { getItem: (name: string) => OfficeTable; getItemOrNullObject?: (name: string) => OfficeTable & { isNullObject?: boolean }; load?: (fields: string | readonly string[]) => void; items?: OfficeTable[] };
    comments?: OfficeComments;
    application?: { calculate: (type: string) => void; calculationState?: string; load?: (fields: string | readonly string[]) => void };
  };
  sync: () => Promise<void>;
};
type OfficeTable = {
  name?: string;
  showHeaders?: boolean;
  showTotals?: boolean;
  isNullObject?: boolean;
  load?: (fields: string | readonly string[]) => void;
  getRange?: () => OfficeRange;
  getDataBodyRange?: () => OfficeRange;
  sort?: { fields?: readonly { key?: number; ascending?: boolean }[]; load?: (fields: string | readonly string[]) => void; apply: (fields: readonly { readonly key: number; readonly ascending: boolean }[]) => void };
  columns?: { getItemAt: (index: number) => { filter?: { criteria?: unknown; load?: (fields: string | readonly string[]) => void; clear: () => void; applyValuesFilter: (values: readonly string[]) => void; applyCustomFilter: (criterion1: string, criterion2?: string, operator?: string) => void } }; items?: unknown[]; load?: (fields: string | readonly string[]) => void };
};
type OfficeRun = (callback: (context: OfficeContext) => Promise<unknown>) => Promise<unknown>;

const MAX_COMMENT_SCAN_ITEMS = 512;
const OPERATION_UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

const nativeExcelRun = (): OfficeRun | undefined => {
  const candidate = (globalThis as Record<string, unknown>).Excel;
  if (candidate === null || typeof candidate !== "object") return undefined;
  const run = (candidate as Record<string, unknown>).run;
  return typeof run === "function" ? run as OfficeRun : undefined;
};

function isScalarValueMatrix(value: unknown): value is ValueMatrix {
  return Array.isArray(value) && value.length > 0 && value.every((row) =>
    Array.isArray(row) && row.length > 0 && row.every((item) =>
      item === null || typeof item === "string" || typeof item === "boolean" || typeof item === "number" && Number.isFinite(item),
    ),
  );
}

function internalObjectName(kind: "Chart" | "Table", operationId: string): string | undefined {
  if (!OPERATION_UUID_PATTERN.test(operationId)) return undefined;
  return `Hermes_${kind}_${operationId.replace(/-/g, "_")}`;
}

export interface NativeWorkbookPlanRuntimeContext {
  readonly compileContext: VerifiedWorkbookCompileContext;
  readonly activeSheetName: string;
  readonly selection?: { readonly sheetName: string; readonly address: string };
}

/** Hydrate only host-owned facts used by compilation and policy. */
export async function readNativeWorkbookPlanRuntimeContext(
  plan?: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
): Promise<NativeWorkbookPlanRuntimeContext> {
  const run = nativeExcelRun();
  if (run === undefined) throw new Error("Office workbook plan runtime is unavailable");
  return run(async (context) => {
    const worksheets = context.workbook.worksheets;
    const active = worksheets.getActiveWorksheet?.();
    const selected = context.workbook.getSelectedRange?.();
    worksheets.load?.("items/name");
    active?.load?.("name");
    selected?.load?.("address");
    const tables = context.workbook.tables;
    tables?.load?.("items/name");
    await context.sync();

    const sheetItems = worksheets.items ?? [];
    const tableItems = tables?.items ?? [];
    for (const table of tableItems) {
      table.load?.("name");
      table.columns?.load?.("items/name");
    }
    const usedRanges = sheetItems.map((item) => {
      const range = item.getUsedRangeOrNullObject?.();
      range?.load?.(["address", "isNullObject"]);
      return { item, range };
    });
    const tableRanges = tableItems.map((table) => {
      const range = table.getRange?.();
      range?.load?.(["address", "rowCount", "columnCount"]);
      return { table, range };
    });
    await context.sync();

    const preconditionManifest = plan === undefined
      ? undefined
      : await captureNativePreconditionManifest(context, plan, sheetItems, tableItems);

    const sheetNames = new Set(sheetItems.map((item) => item.name).filter((name): name is string => typeof name === "string" && name.length > 0));
    const tablesByName = new Map<string, { canonicalName: string; columnCount: number; rowCount?: number; sheetName?: string }>();
    for (const table of tableItems) {
      const name = (table as Record<string, unknown>).name;
      if (typeof name === "string" && name.length > 0) {
        const tableRange = tableRanges.find((candidate) => candidate.table === table)?.range;
        const tableAddress = tableRange?.address;
        const bang = typeof tableAddress === "string" ? tableAddress.indexOf("!") : -1;
        tablesByName.set(name, {
          canonicalName: name,
          columnCount: typeof tableRange?.columnCount === "number"
            ? tableRange.columnCount
            : table.columns?.items?.length ?? 0,
          ...(typeof tableRange?.rowCount === "number" ? { rowCount: tableRange.rowCount } : {}),
          ...(bang > 0 && typeof tableAddress === "string"
            ? { sheetName: tableAddress.slice(0, bang).replace(/^'|'$/g, "") }
            : {}),
        });
      }
    }
    const chartSources = new Map<string, ReadonlyMap<string, { areaCount: number }>>();
    for (const { item, range } of usedRanges) {
      const name = item.name;
      const address = range !== undefined && range.isNullObject !== true ? (range as Record<string, unknown>).address : undefined;
      if (typeof name !== "string" || typeof address !== "string") continue;
      const byAddress = new Map<string, { areaCount: number }>();
      byAddress.set(address, { areaCount: 1 });
      const bang = address.indexOf("!");
      if (bang >= 0) byAddress.set(address.slice(bang + 1), { areaCount: 1 });
      chartSources.set(name, byAddress);
    }
    const supportedPrimitiveNames = new Set(
      workbookPlanContract.operationDefinitions.flatMap((definition) => [...definition.compilerTargets]),
    );
    const supportedCapabilities = new Set(
      workbookPlanContract.operationDefinitions.flatMap((definition) => definition.capabilityRequirements.map(({ capability }) => capability)),
    );
    const selectedAddress = selected?.address;
    const selectionBang = typeof selectedAddress === "string" ? selectedAddress.indexOf("!") : -1;
    return {
      activeSheetName: typeof active?.name === "string" ? active.name : [...sheetNames][0] ?? "",
      ...(selectionBang > 0 && typeof selectedAddress === "string"
        ? {
            selection: {
              sheetName: selectedAddress.slice(0, selectionBang).replace(/^'|'$/g, ""),
              address: selectedAddress.slice(selectionBang + 1),
            },
          }
        : {}),
      compileContext: {
        sheetNames,
        tablesByName,
        supportedPrimitiveNames,
        supportedCapabilities,
        chartSourceResolutionsBySheet: chartSources,
        ...(preconditionManifest === undefined ? {} : { preconditionManifest }),
      },
    };
  }) as Promise<NativeWorkbookPlanRuntimeContext>;
}

function captureNativePreconditionManifest(
  context: OfficeContext,
  plan: Extract<WorkbookPlan, { readonly outcome: "execute" }>,
  sheetItems: readonly OfficeWorksheet[],
  tableItems: readonly OfficeTable[],
): Promise<WorkbookPlanPreconditionManifest> {
  return (async () => {
    const worksheets = context.workbook.worksheets;
    const currentSheets = new Set(sheetItems.map((sheet) => sheet.name).filter((name): name is string => typeof name === "string"));
    const sheetNames = new Set<string>(currentSheets);
    const tableNames = new Set<string>(tableItems.map((table) => (table as Record<string, unknown>).name).filter((name): name is string => typeof name === "string"));
    const rangeKeys = new Map<string, { readonly sheet: string; readonly address: string; readonly source: boolean }>();
    const chartSourceKeys = new Map<string, { readonly sheet: string; readonly address: string }>();
    const chartKeys = new Map<string, { readonly sheet: string; readonly title: string }>();
    const commentKeys = new Map<string, { readonly sheet: string; readonly address: string }>();
    const tableTargets = new Set<string>();
    const addRange = (sheet: string, address: string, source = false): void => {
      const key = `${sheet}\u0000${address}`;
      const previous = rangeKeys.get(key);
      rangeKeys.set(key, { sheet, address, source: source || previous?.source === true });
    };
    for (const operation of plan.operations) {
      switch (operation.kind) {
        case "create_sheet": sheetNames.add(operation.name); break;
        case "rename_sheet": sheetNames.add(operation.oldName); sheetNames.add(operation.newName); break;
        case "copy_sheet": sheetNames.add(operation.sourceName); sheetNames.add(operation.newName); break;
        case "sort_table":
        case "filter_table":
        case "reset_table_filter": tableNames.add(operation.tableName); tableTargets.add(operation.tableName); break;
        case "fill_formula":
          sheetNames.add(operation.sheet);
          addRange(operation.sheet, operation.sourceAddress, true);
          addRange(operation.sheet, operation.targetAddress);
          break;
        case "create_chart":
          sheetNames.add(operation.sheet);
          addRange(operation.sheet, operation.dataRange);
          chartSourceKeys.set(`${operation.sheet}\u0000${operation.dataRange}`, { sheet: operation.sheet, address: operation.dataRange });
          chartKeys.set(`${operation.sheet}\u0000${operation.title}`, { sheet: operation.sheet, title: operation.title });
          break;
        case "add_comment":
          sheetNames.add(operation.sheet);
          addRange(operation.sheet, operation.address);
          commentKeys.set(`${operation.sheet}\u0000${operation.address}`, { sheet: operation.sheet, address: operation.address });
          break;
        case "calculate":
        case "freeze_panes":
          if (operation.kind === "freeze_panes") sheetNames.add(operation.sheet);
          break;
        default:
          sheetNames.add(operation.sheet);
          addRange(operation.sheet, operation.address);
          break;
      }
    }
    for (const rule of plan.verification) {
      switch (rule.kind) {
        case "range_values_equal":
        case "range_formulas_equal":
        case "range_formula_pattern":
        case "range_number_format_equal":
        case "sentinel_unchanged":
          sheetNames.add(rule.sheet); addRange(rule.sheet, rule.address); break;
        case "chart_exists":
          sheetNames.add(rule.sheet); chartKeys.set(`${rule.sheet}\u0000${rule.title}`, { sheet: rule.sheet, title: rule.title }); break;
        case "comment_exists":
          sheetNames.add(rule.sheet); addRange(rule.sheet, rule.address); commentKeys.set(`${rule.sheet}\u0000${rule.address}`, { sheet: rule.sheet, address: rule.address }); break;
        case "table_exists":
        case "table_row_order":
        case "table_filter_state":
          tableNames.add(rule.tableName); break;
        case "sheet_exists": sheetNames.add(rule.name); break;
        case "calculation_completed": break;
      }
    }
    if (worksheets.getItemOrNullObject === undefined) throw new Error("Worksheet existence capture unavailable");
    const sheetTargets = [...sheetNames].map((name) => {
      const target = worksheets.getItemOrNullObject!(name);
      if (target === undefined || typeof target.load !== "function") throw new Error("Worksheet existence capture unavailable");
      target.load("isNullObject");
      return { name, target };
    });
    const tables = context.workbook.tables;
    if (tables?.getItemOrNullObject === undefined) throw new Error("Table existence capture unavailable");
    const tableTargetsQueued = [...tableNames].map((name) => {
      const target = tables.getItemOrNullObject!(name);
      if (target === undefined || typeof target.load !== "function") throw new Error("Table existence capture unavailable");
      target.load("isNullObject");
      return { name, target };
    });
    const rangeTargets = [...rangeKeys.values()].map((expected) => {
      const target = worksheets.getItem(expected.sheet).getRange(expected.address);
      if (target === undefined || typeof target.load !== "function") throw new Error("Range capture unavailable");
      target.load(["rowCount", "columnCount", "address", "values", "formulas", "numberFormat"]);
      return { ...expected, target };
    });
    const chartSourceTargets = [...chartSourceKeys.values()].map((expected) => {
      const target = worksheets.getItem(expected.sheet).getRanges?.(expected.address);
      if (target === undefined || typeof target.load !== "function") throw new Error("Chart source capture unavailable");
      target.load(["address", "areaCount"]);
      return { ...expected, target };
    });
    const chartTargets = [...chartKeys.values()].map((expected) => {
      const charts = worksheets.getItem(expected.sheet).charts;
      if (charts === undefined || typeof charts.load !== "function") throw new Error("Chart capture unavailable");
      charts.load("items/title/text");
      return { ...expected, target: charts };
    });
    const commentTargets = new Map<string, OfficeCommentRequest>();
    for (const expected of commentKeys.values()) {
      if (commentTargets.has(expected.sheet)) continue;
      const targetSheet = worksheets.getItem(expected.sheet);
      const comments = targetSheet.comments;
      if (comments === undefined) throw new Error("Comment capture unavailable");
      commentTargets.set(expected.sheet, { sheet: expected.sheet, comments });
    }
    await context.sync();
    const tableRangesQueued = tableTargetsQueued.map(({ name, target }) => {
      if (target.isNullObject === true) return { name, target, range: undefined };
      const range = target.getRange?.();
      if (range === undefined || typeof range.load !== "function") throw new Error("Table dimension capture unavailable");
      range.load(["rowCount", "columnCount", "address"]);
      return { name, target, range };
    });
    await context.sync();
    const commentScans: OfficeCommentScans = commentTargets.size === 0
      ? new Map<string, OfficeCommentScan>()
      : await scanCommentCollections(context, [...commentTargets.values()]);
    const commentReadbacks = [...commentKeys.values()].map(({ sheet, address }) => {
      const scan = commentScans.get(sheet);
      if (scan === undefined) throw new Error("Comment capture readback unavailable");
      return { sheet, address, pending: scan.candidates };
    });
    const sheets = sheetTargets.map(({ name, target }) => ({ name, exists: target.isNullObject !== true }));
    const tablesManifest = tableRangesQueued.map(({ name, target, range }) => {
      const address = range?.address;
      const bang = typeof address === "string" ? address.indexOf("!") : -1;
      return {
        name,
        exists: target.isNullObject !== true,
        ...(typeof range?.columnCount === "number" ? { columnCount: range.columnCount } : {}),
        ...(typeof range?.rowCount === "number" ? { rowCount: range.rowCount } : {}),
        ...(bang > 0 && typeof address === "string" ? { sheetName: address.slice(0, bang).replace(/^'|'$/g, "") } : {}),
        ...(typeof address === "string" ? { address } : {}),
      };
    });
    const ranges: WorkbookPlanPreconditionRange[] = rangeTargets.map(({ sheet, address, target, source }) => {
      const values = target.values as unknown;
      const formulas = target.formulas as unknown;
      const numberFormat = target.numberFormat as unknown;
      const firstValue = Array.isArray(values) && Array.isArray(values[0]) ? values[0][0] : undefined;
      return {
        sheet,
        address,
        ...(typeof target.address === "string" ? { resolvedAddress: target.address } : {}),
        rowCount: target.rowCount ?? 0,
        columnCount: target.columnCount ?? 0,
        sentinel: firstValue as ScalarValue,
        ...(isScalarValueMatrix(values) ? { values: values.map((row) => [...row]) } : {}),
        ...(source && Array.isArray(formulas) ? { sourceFormulas: formulas as readonly (readonly string[])[] } : {}),
        ...(Array.isArray(numberFormat) && Array.isArray(numberFormat[0]) && typeof numberFormat[0][0] === "string"
          ? { numberFormat: numberFormat[0][0] as string }
          : {}),
      };
    });
    return {
      version: 1,
      sheets,
      tables: tablesManifest,
      ranges,
      chartSources: chartSourceTargets.map(({ sheet, address, target }) => ({
        sheet,
        address,
        areaCount: target.areaCount ?? 0,
        ...(typeof target.address === "string" ? { resolvedAddress: target.address } : {}),
      })),
      charts: chartTargets.map(({ sheet, title, target }) => {
        if (!Array.isArray(target.items)) throw new Error("Chart capture readback unavailable");
        return { sheet, title, exists: target.items.some((chart) => chart.title?.text === title) };
      }),
      comments: commentReadbacks.map(({ sheet, address, pending }) => {
        const expected = normalizedCommentAddress(address);
        const match = pending.find(({ location }) => typeof location.address === "string" && normalizedCommentAddress(location.address) === expected);
        return {
        sheet,
        address,
        exists: match !== undefined,
        ...(match !== undefined && typeof match.comment.content === "string" ? { text: match.comment.content } : {}),
        };
      }),
    };
  })();
}

/**
 * Native, route-neutral Office.js adapter. A group is queued in exactly one
 * Excel.run; bounded collection/location phases may use their own sync calls,
 * but the adapter never re-enters Excel.run per primitive or falls back to
 * dispatching primitive tool calls through the compatibility executor.
 */
export function createNativeWorkbookPlanAdapter(): WorkbookPlanExecutionAdapter & {
  readonly verifyPlan: (
    rules: readonly VerificationRule[],
    signal?: AbortSignal,
  ) => Promise<WorkbookPlanVerificationReport>;
} {
  const operationIdentities = new Map<string, OperationIdentityBefore>();
  const executeGroup = async (
    group: WorkbookPlanEffectGroup,
    _groupIndex: number,
    signal?: AbortSignal,
  ): Promise<WorkbookPlanGroupExecutionResult> => {
    const run = nativeExcelRun();
    if (run === undefined || isOfficeCancellationRequested(signal)) return { status: "sync_failed", failureCode: "sync_failed" };
    try {
      await run(async (context) => {
        if (isOfficeCancellationRequested(signal)) throw new Error("Office mutation handoff cancelled");
        const commentRequests = commentRequestsForOperations(context, group.operations);
        const commentScans: OfficeCommentScans = commentRequests.length === 0
          ? new Map<string, OfficeCommentScan>()
          : await scanCommentCollections(context, commentRequests);
        const commentCache = commentCacheForOperations(group.operations, commentScans);
        if (isOfficeCancellationRequested(signal)) throw new Error("Office mutation handoff cancelled");
        await captureOperationIdentities(context, group.operations, operationIdentities, commentScans);
        await queueGroup(context, group, commentCache, operationIdentities);
        if (isOfficeCancellationRequested(signal)) throw new Error("Office mutation handoff cancelled");
        await context.sync();
      });
      return { status: "committed" };
    } catch {
      return { status: "sync_failed", failureCode: "sync_failed" };
    }
  };

  return {
  releaseExecutionState(operations) {
    for (const operation of operations) {
      if (operationIdentities.get(operation.operationId)?.operation === operation) operationIdentities.delete(operation.operationId);
    }
  },
  async executeFirstGroup(
      gate: WorkbookPlanFirstGroupContext,
      signal?: AbortSignal,
    ): Promise<WorkbookPlanGroupExecutionResult> {
      const run = nativeExcelRun();
      if (run === undefined || gate.preconditionManifest === undefined || signal !== undefined && signal.aborted) {
        return { status: "sync_failed", failureCode: "sync_failed" };
      }
      try {
        return await run(async (context) => {
          return gate.executeAfterFresh("fresh", async (group) => {
            const preconditions = queueManifestPreconditions(context, gate.preconditionManifest!);
            const commentRequests = [
              ...commentRequestsForPreconditions(preconditions),
              ...commentRequestsForOperations(context, group.operations),
            ];
            const commentScans = await scanCommentCollections(context, commentRequests);
            attachCommentPreconditionCandidates(preconditions, commentScans);
            if (isOfficeCancellationRequested(signal)) return { status: "sync_failed", failureCode: "sync_failed" };
            if (!preconditionsFresh(preconditions)) return { status: "stale_precondition" };
            if (isOfficeCancellationRequested(signal)) return { status: "sync_failed", failureCode: "sync_failed" };
            await captureOperationIdentities(context, group.operations, operationIdentities, commentScans);
            await queueGroup(context, group, commentCacheForOperations(group.operations, commentScans), operationIdentities);
            if (isOfficeCancellationRequested(signal)) throw new Error("Office mutation handoff cancelled");
            await context.sync();
            return { status: "committed" };
          });
        }) as WorkbookPlanGroupExecutionResult;
      } catch {
        return { status: "sync_failed", failureCode: "sync_failed" };
      }
    },
    executeGroup,
    async reconcileFailedGroup(group) {
      const run = nativeExcelRun();
      if (run === undefined) return group.operations.map(() => "unknown_within_batch" as const);
      try {
        return await run(async (context) => {
          const pending = group.operations.map((operation) => ({
            operation,
            request: (() => {
              const request = queueReconciliationRead(context, operation);
              return request === undefined ? undefined : {
                ...request,
                identityBefore: (() => {
                  const identity = operationIdentities.get(operation.operationId);
                  return identity?.operation === operation ? identity : undefined;
                })(),
              };
            })(),
          }));
          const commentRequests = pending.flatMap(({ request }) => request?.kind === "comments" && request.commentSheet !== undefined
            ? [{ sheet: request.commentSheet, comments: request.target as OfficeComments }]
            : []);
          const commentScans = await scanCommentCollections(context, commentRequests);
          attachCommentReadCandidates(pending.map(({ request }) => request), commentScans);
          return pending.map(({ operation, request }) => reconcileOperation(operation, request));
        }) as readonly ("committed" | "not_committed" | "unknown_within_batch")[];
      } catch {
        return group.operations.map(() => "unknown_within_batch" as const);
      }
    },
    async verifyMandatoryPostconditions(proofs, operations, signal) {
      if (isOfficeCancellationRequested(signal)) return false;
      const run = nativeExcelRun();
      if (run === undefined) return false;
      try {
        return await run(async (context) => {
          const pending = proofs.map((proof) => {
            const operation = operations.find((candidate) => candidate.operationIndex === proof.operationIndex);
            if (operation === undefined || proof.proof === undefined) return { operation, request: undefined };
            const request = queueReconciliationRead(context, operation);
            if (request === undefined) return { operation, request };
            return {
              operation,
              request: {
                ...request,
                ...(proof.proof.expectedFormulas === undefined ? {} : { expectedFormulas: proof.proof.expectedFormulas }),
                identityBefore: (() => {
                  const identity = operationIdentities.get(operation.operationId);
                  return identity?.operation === operation ? identity : undefined;
                })(),
              },
            };
          });
          const commentRequests = pending.flatMap(({ request }) => request?.kind === "comments" && request.commentSheet !== undefined
            ? [{ sheet: request.commentSheet, comments: request.target as OfficeComments }]
            : []);
          const commentScans = await scanCommentCollections(context, commentRequests);
          attachCommentReadCandidates(pending.map(({ request }) => request), commentScans);
          await context.sync();
          if (isOfficeCancellationRequested(signal)) return false;
          return pending.every(({ operation, request }) => operation !== undefined && reconcileOperation(operation, request) === "committed");
        }) as boolean;
      } catch { return false; }
    },
    async verifyPlan(rules, signal) {
      if (signal !== undefined && signal.aborted) {
        return failedVerification();
      }
      const prepared = prepareWorkbookPlanVerificationRules(JSON.stringify(rules));
      if (prepared === null) return failedVerification();
      const run = nativeExcelRun();
      if (run === undefined) return failedVerification();
      const readbacks = new Map<number, string>();
      try {
        await run(async (context) => {
          const pending = rules.map((rule, index) => ({ rule, index, request: queueRead(context, rule) }));
          const commentRequests = pending.flatMap(({ request }) => request?.kind === "comments" && request.commentSheet !== undefined
            ? [{ sheet: request.commentSheet, comments: request.target as OfficeComments }]
            : []);
          const commentScans = await scanCommentCollections(context, commentRequests);
          attachCommentReadCandidates(pending.map(({ request }) => request), commentScans);
          await context.sync();
          for (const item of pending) {
            const snapshot = readSnapshot(item.rule, item.request);
            if (snapshot !== undefined) {
              readbacks.set(item.index, JSON.stringify({ status: "supported", snapshot }));
            } else {
              readbacks.set(item.index, JSON.stringify({ status: "unsupported", reason: "native_readback_unavailable" }));
            }
          }
        });
      } catch {
        return failedVerification();
      }
      if (signal !== undefined && signal.aborted) {
        return failedVerification();
      }
      return verifyWorkbookPlan(prepared, {
        read: () => undefined,
        readAt: (index) => readbacks.get(index),
      });
    },
  };
}

async function queueGroup(
  context: OfficeContext,
  group: WorkbookPlanEffectGroup,
  commentCache: Map<string, OfficeComment | undefined>,
  identities: Map<string, OperationIdentityBefore>,
): Promise<void> {
  for (const operation of group.operations) await queueOperation(context, operation, commentCache, identities);
}

async function queueOperation(
  context: OfficeContext,
  operation: WorkbookPlanEffectOperation,
  commentCache: Map<string, OfficeComment | undefined>,
  identities: Map<string, OperationIdentityBefore>,
): Promise<void> {
  const args = operation.args;
  const workbook = context.workbook;
  switch (operation.kind) {
    case "write_values": {
      const range = sheet(context, stringArg(args, "sheet")).getRange(stringArg(args, "address"));
      range.values = args.values;
      return;
    }
    case "write_formula": {
      const range = sheet(context, stringArg(args, "sheet")).getRange(stringArg(args, "address"));
      range.formulas = args.values;
      return;
    }
    case "format_range": {
      const address = stringArg(args, "address");
      const range = sheet(context, stringArg(args, "sheet")).getRange(address);
      applyFormat(range, recordArg(args, "format"), address);
      return;
    }
    case "fill_formula": {
      const targetSheet = sheet(context, stringArg(args, "sheet"));
      const source = targetSheet.getRange(stringArg(args, "sourceAddress"));
      const target = targetSheet.getRange(stringArg(args, "targetAddress"));
      if (typeof source.autoFill !== "function") throw new Error("Office autoFill unavailable");
      source.autoFill(target);
      return;
    }
    case "create_table": {
      const targetSheet = sheet(context, stringArg(args, "sheet"));
      if (targetSheet.tables === undefined) throw new Error("Office tables unavailable");
      const objectName = internalObjectName("Table", operation.operationId);
      if (objectName === undefined) throw new Error("Invalid trusted table operation identity");
      const table = targetSheet.tables.add(targetSheet.getRange(stringArg(args, "address")), args.hasHeaders === true);
      table.name = objectName;
      return;
    }
    case "sort_table":
    case "filter_table":
    case "reset_table_filter":
      queueTableOperation(workbook, operation);
      return;
    case "create_chart": {
      const target = sheet(context, stringArg(args, "sheet"));
      if (target.charts === undefined) throw new Error("Office charts unavailable");
      const source = identities.get(operation.operationId)?.sourceRange;
      if (source === undefined) throw new Error("Trusted chart source identity unavailable");
      const chart = target.charts.add(stringArg(args, "chartType"), source, "Auto");
      const objectName = internalObjectName("Chart", operation.operationId);
      if (objectName === undefined) throw new Error("Invalid trusted chart operation identity");
      chart.name = objectName;
      if (chart.title === undefined) throw new Error("Office chart title unavailable");
      chart.title.text = stringArg(args, "title");
      return;
    }
    case "set_validation_list": {
      const range = sheet(context, stringArg(args, "sheet")).getRange(stringArg(args, "address"));
      const validation = (range as Record<string, unknown>).dataValidation as Record<string, unknown> | undefined;
      if (validation === undefined) throw new Error("Office data validation unavailable");
      validation.rule = { list: { inCellDropDown: true, source: stringArg(args, "source") } };
      if (typeof args.promptTitle === "string" || typeof args.promptMessage === "string") {
        validation.prompt = {
          showPrompt: true,
          title: typeof args.promptTitle === "string" ? args.promptTitle : "",
          message: typeof args.promptMessage === "string" ? args.promptMessage : "",
        };
      }
      return;
    }
    case "conditional_format": {
      const range = sheet(context, stringArg(args, "sheet")).getRange(stringArg(args, "address"));
      const formats = (range as Record<string, unknown>).conditionalFormats as { add: (type: string) => Record<string, unknown> } | undefined;
      if (formats === undefined) throw new Error("Office conditional formats unavailable");
      const ruleType = normalizeConditionalType(stringArg(args, "ruleType"));
      const conditional = formats.add(ruleType);
      if (ruleType === "CellValue") {
        const cellValue = conditional.cellValue as Record<string, unknown> | undefined;
        if (cellValue === undefined) throw new Error("Office CellValue conditional format unavailable");
        cellValue.rule = conditionalCellValueRule(args);
        applyConditionalFormat(cellValue.format, args);
      } else {
        const textComparison = conditional.textComparison as Record<string, unknown> | undefined;
        if (textComparison === undefined) throw new Error("Office TextComparison conditional format unavailable");
        textComparison.rule = conditionalTextComparisonRule(args);
        applyConditionalFormat(textComparison.format, args);
      }
      return;
    }
    case "freeze_panes": {
      const panes = (sheet(context, stringArg(args, "sheet")) as unknown as Record<string, unknown>).freezePanes as {
        freezeRows: (n: number) => void;
        freezeColumns: (n: number) => void;
        unfreeze?: () => void;
      } | undefined;
      if (panes === undefined) throw new Error("Office freeze panes unavailable");
      const rows = numberArg(args, "rows");
      const columns = numberArg(args, "columns");
      if (rows === 0 && columns === 0) {
        if (typeof panes.unfreeze !== "function") throw new Error("Office pane unfreeze unavailable");
        panes.unfreeze();
      } else {
        if (rows > 0) panes.freezeRows(rows);
        if (columns > 0) panes.freezeColumns(columns);
      }
      return;
    }
    case "create_sheet":
      if (workbook.worksheets.add === undefined) throw new Error("Office worksheet creation unavailable");
      workbook.worksheets.add(stringArg(args, "name"));
      return;
    case "rename_sheet": {
      const target = workbook.worksheets.getItem(stringArg(args, "oldName"));
      target.name = stringArg(args, "newName");
      return;
    }
    case "copy_sheet": {
      const target = workbook.worksheets.getItem(stringArg(args, "sourceName"));
      if (typeof target.copy !== "function") throw new Error("Office sheet copy unavailable");
      const copied = target.copy("End");
      copied.name = stringArg(args, "newName");
      return;
    }
    case "add_comment": {
      const sheetName = stringArg(args, "sheet");
      const address = stringArg(args, "address");
      await queueAddOrUpdateComment(context, sheetName, address, stringArg(args, "text"), commentCache);
      return;
    }
    case "calculate":
      if (workbook.application === undefined) throw new Error("Office calculation unavailable");
      workbook.application.calculate("Full");
      return;
    default: {
      const exhaustive: never = operation.kind;
      throw new Error(`Unsupported compiled operation ${exhaustive}`);
    }
  }
}

function sheet(context: OfficeContext, name: string): OfficeWorksheet {
  return context.workbook.worksheets.getItem(name);
}

function queueTableOperation(workbook: OfficeContext["workbook"], operation: WorkbookPlanEffectOperation): void {
  const tables = workbook.tables;
  if (tables === undefined) throw new Error("Office tables unavailable");
  const table = tables.getItem(stringArg(operation.args, "tableName"));
  const column = numberArg(operation.args, "columnIndex");
  if (operation.kind === "sort_table") {
    if (table.sort === undefined) throw new Error("Office table sorting unavailable");
    table.sort.apply([{ key: column, ascending: operation.args.ascending !== false }]);
    return;
  }
  const filter = table.columns?.getItemAt(column).filter;
  if (filter === undefined) throw new Error("Office table filters unavailable");
  if (operation.kind === "reset_table_filter") {
    filter.clear();
  } else if (operation.args.values !== undefined) {
    filter.applyValuesFilter(operation.args.values as readonly string[]);
  } else {
    const criterion1 = typeof operation.args.criterion1 === "string" ? operation.args.criterion1 : undefined;
    const criterion2 = typeof operation.args.criterion2 === "string" ? operation.args.criterion2 : undefined;
    const operator = typeof operation.args.operator === "string" ? operation.args.operator : undefined;
    if (criterion1 === undefined) filter.clear();
    else if (operator !== undefined) filter.applyCustomFilter(criterion1, criterion2, operator);
    else if (criterion2 !== undefined) filter.applyCustomFilter(criterion1, criterion2);
    else filter.applyCustomFilter(criterion1);
  }
}

function applyFormat(range: OfficeRange, format: Record<string, unknown>, address: string): void {
  const target = range as Record<string, unknown>;
  const officeFormat = target.format as Record<string, unknown> | undefined;
  if (officeFormat === undefined) throw new Error("Office range format unavailable");
  const font = officeFormat.font as Record<string, unknown> | undefined;
  if (font !== undefined) {
    if (typeof format.bold === "boolean") font.bold = format.bold;
    if (typeof format.italic === "boolean") font.italic = format.italic;
    if (typeof format.fontColor === "string") font.color = format.fontColor;
    if (typeof format.fontName === "string") font.name = format.fontName;
    if (typeof format.fontSize === "number") font.size = format.fontSize;
    if (typeof format.underline === "string") font.underline = format.underline;
  }
  if (typeof format.fillColor === "string") {
    const fill = officeFormat.fill as Record<string, unknown> | undefined;
    if (fill === undefined) throw new Error("Office range fill unavailable");
    fill.color = format.fillColor;
  }
  for (const key of ["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation"] as const) {
    if (format[key] !== undefined) officeFormat[key] = format[key];
  }
  if (typeof format.numberFormat === "string") target.numberFormat = dimensionedNumberFormat(address, format.numberFormat);
  if (format.borders === true) {
    const borders = officeFormat.borders as { getItem: (name: string) => Record<string, unknown> } | undefined;
    if (borders === undefined) throw new Error("Office range borders unavailable");
    for (const name of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
      const border = borders.getItem(name);
      border.style = "Continuous";
      if (typeof format.borderColor === "string") border.color = format.borderColor;
      if (typeof format.borderWeight === "string") border.weight = format.borderWeight;
    }
  }
  if (format.columnAutofit === true) {
    const autofit = officeFormat.autofitColumns;
    if (typeof autofit !== "function") throw new Error("Office column autofit unavailable");
    autofit.call(officeFormat);
  }
  if (format.rowAutofit === true) {
    const autofit = officeFormat.autofitRows;
    if (typeof autofit !== "function") throw new Error("Office row autofit unavailable");
    autofit.call(officeFormat);
  }
}

function dimensionedNumberFormat(address: string, value: string): string[][] {
  const local = address.split("!").at(-1)?.replace(/\$/g, "") ?? address;
  const [first, last = first] = local.split(":");
  const parseCell = (cell: string): { row: number; column: number } => {
    const match = /^([A-Z]+)([0-9]+)$/i.exec(cell);
    if (match === null) throw new Error(`Unsupported numberFormat range ${address}`);
    let column = 0;
    for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
    return { row: Number(match[2]), column };
  };
  const start = parseCell(first!);
  const end = parseCell(last!);
  const rows = end.row - start.row + 1;
  const columns = end.column - start.column + 1;
  if (rows <= 0 || columns <= 0) throw new Error(`Invalid numberFormat range ${address}`);
  return Array.from({ length: rows }, () => Array(columns).fill(value));
}

function normalizeConditionalType(value: string): "CellValue" | "ContainsText" {
  if (value === "cellValue" || value === "CellValue") return "CellValue";
  if (value === "containsText" || value === "ContainsText") return "ContainsText";
  throw new Error(`Unsupported conditional format type ${value}`);
}

function normalizeCellValueOperator(value: string): string {
  const operators: Record<string, string> = {
    between: "Between",
    notBetween: "NotBetween",
    equalTo: "EqualTo",
    notEqualTo: "NotEqualTo",
    greaterThan: "GreaterThan",
    lessThan: "LessThan",
    greaterThanOrEqual: "GreaterThanOrEqual",
    lessThanOrEqual: "LessThanOrEqual",
  };
  const normalized = operators[value] ?? value;
  if (!Object.values(operators).includes(normalized)) throw new Error(`Unsupported CellValue operator ${value}`);
  return normalized;
}

function normalizeTextComparisonOperator(value: string): string {
  const operators: Record<string, string> = {
    contains: "Contains",
    notContains: "NotContains",
    beginsWith: "BeginsWith",
    endsWith: "EndsWith",
  };
  const normalized = operators[value] ?? value;
  if (!Object.values(operators).includes(normalized)) throw new Error(`Unsupported TextComparison operator ${value}`);
  return normalized;
}

function conditionalCellValueRule(args: Readonly<Record<string, unknown>>): Record<string, unknown> {
  const operator = typeof args.operator === "string"
    ? normalizeCellValueOperator(args.operator)
    : "GreaterThan";
  const formula1 = stringArg(args, "formula1");
  return {
    operator,
    formula1,
    ...(typeof args.formula2 === "string" ? { formula2: args.formula2 } : {}),
  };
}

function conditionalTextComparisonRule(args: Readonly<Record<string, unknown>>): Record<string, unknown> {
  return {
    operator: typeof args.operator === "string"
      ? normalizeTextComparisonOperator(args.operator)
      : "Contains",
    text: stringArg(args, "text"),
  };
}

function applyConditionalFormat(formatValue: unknown, args: Readonly<Record<string, unknown>>): void {
  const format = formatValue as Record<string, unknown> | undefined;
  if (format === undefined) throw new Error("Office conditional format differential unavailable");
  const fill = format.fill as Record<string, unknown> | undefined;
  const font = format.font as Record<string, unknown> | undefined;
  if (typeof args.fillColor === "string") {
    if (fill === undefined) throw new Error("Office conditional fill unavailable");
    fill.color = args.fillColor;
  }
  if (typeof args.fontColor === "string") {
    if (font === undefined) throw new Error("Office conditional font unavailable");
    font.color = args.fontColor;
  }
}

function normalizedCommentAddress(address: string): string {
  return (address.split("!").at(-1) ?? address).replace(/\$/g, "").toUpperCase();
}

async function scanCommentCollections(
  context: OfficeContext,
  requests: readonly OfficeCommentRequest[],
): Promise<OfficeCommentScans> {
  if (requests.length === 0) {
    await context.sync();
    return new Map<string, OfficeCommentScan>();
  }
  // Office.js may return distinct worksheet/comment proxies for the same logical
  // sheet. The first request is the canonical collection for this bounded scan;
  // deduplicate by the trusted sheet key rather than ClientObject identity.
  const bySheet = new Map<string, OfficeCommentRequest>();
  for (const request of requests) {
    if (typeof request.comments.getCount !== "function") throw new Error("Comment count unavailable");
    if (!bySheet.has(request.sheet)) bySheet.set(request.sheet, request);
  }

  const countResults = [...bySheet.values()].map((request) => ({
    request,
    result: request.comments.getCount(),
  }));
  await context.sync();
  const bounded = countResults.map(({ request, result }) => {
    const count = result.value;
    if (!Number.isSafeInteger(count) || count < 0 || count > MAX_COMMENT_SCAN_ITEMS) {
      throw new Error("Comment collection count exceeds conservative scan bound");
    }
    return { request, count };
  });

  for (const { request } of bounded) request.comments.load("items");
  await context.sync();
  const pending = bounded.map(({ request, count }) => {
    const items = request.comments.items;
    if (!Array.isArray(items) || !Number.isSafeInteger(items.length) || items.length > MAX_COMMENT_SCAN_ITEMS || items.length !== count) {
      throw new Error("Comment collection items are unavailable or inconsistent with bounded count");
    }
    for (const comment of items) {
      if (comment === null || typeof comment !== "object" || typeof comment.getLocation !== "function" || typeof comment.load !== "function") {
        throw new Error("Comment item readback unavailable");
      }
    }
    return {
      request,
      count,
      items,
      candidates: items.map((comment) => {
        const location = comment.getLocation();
        if (location === null || typeof location !== "object" || typeof location.load !== "function") {
          throw new Error("Comment location readback unavailable");
        }
        location.load("address");
        const loadComment = comment.load;
        if (typeof loadComment !== "function") throw new Error("Comment content readback unavailable");
        loadComment.call(comment, "content");
        return { comment, location };
      }),
    };
  });
  await context.sync();

  const result = new Map<string, OfficeCommentScan>();
  for (const item of pending) {
    const index = new Map<string, OfficeComment>();
    for (const candidate of item.candidates) {
      if (typeof candidate.location.address !== "string" || typeof candidate.comment.content !== "string") {
        throw new Error("Comment location or content readback unavailable");
      }
      const key = normalizedCommentAddress(candidate.location.address);
      if (!index.has(key)) index.set(key, candidate.comment);
    }
    result.set(item.request.sheet, {
      comments: item.request.comments,
      count: item.count,
      items: item.items,
      candidates: item.candidates,
      index,
    });
  }
  return result;
}

function commentRequestsForOperations(
  context: OfficeContext,
  operations: readonly WorkbookPlanEffectOperation[],
): readonly OfficeCommentRequest[] {
  const sheets = new Set<string>();
  for (const operation of operations) {
    if (operation.kind === "add_comment") sheets.add(stringArg(operation.args, "sheet"));
  }
  return [...sheets].map((sheetName) => {
    const comments = sheet(context, sheetName).comments;
    if (comments === undefined) throw new Error("Office comments unavailable");
    return { sheet: sheetName, comments };
  });
}

function commentCacheForOperations(
  operations: readonly WorkbookPlanEffectOperation[],
  scans: OfficeCommentScans,
): Map<string, OfficeComment | undefined> {
  const cache = new Map<string, OfficeComment | undefined>();
  for (const operation of operations) {
    if (operation.kind !== "add_comment") continue;
    const sheetName = stringArg(operation.args, "sheet");
    const address = stringArg(operation.args, "address");
    const scan = scans.get(sheetName);
    if (scan === undefined) throw new Error("Comment scan cache unavailable");
    cache.set(`${sheetName}!${normalizedCommentAddress(address)}`, scan.index.get(normalizedCommentAddress(address)));
  }
  return cache;
}

function commentRequestsForPreconditions(
  preconditions: readonly OfficePrecondition[],
): readonly OfficeCommentRequest[] {
  return preconditions.flatMap((precondition) => {
    if (precondition.kind !== "comments" || precondition.commentSheet === undefined) return [];
    return [{ sheet: precondition.commentSheet, comments: precondition.target as unknown as OfficeComments }];
  });
}

function attachCommentPreconditionCandidates(
  preconditions: readonly OfficePrecondition[],
  scans: OfficeCommentScans,
): void {
  for (const precondition of preconditions) {
    if (precondition.kind !== "comments") continue;
    if (precondition.commentSheet === undefined) throw new Error("Comment precondition sheet unavailable");
    const scan = scans.get(precondition.commentSheet);
    if (scan === undefined) throw new Error("Comment precondition scan unavailable");
    (precondition as { commentCandidates?: readonly OfficeCommentCandidate[] }).commentCandidates = scan.candidates;
  }
}

function attachCommentReadCandidates(
  requests: readonly (OfficeReadRequest | undefined)[],
  scans: OfficeCommentScans,
): void {
  for (const request of requests) {
    if (request?.kind !== "comments") continue;
    if (request.commentSheet === undefined) throw new Error("Comment readback sheet unavailable");
    const scan = scans.get(request.commentSheet);
    if (scan === undefined) throw new Error("Comment readback scan unavailable");
    (request as { commentCandidates?: readonly OfficeCommentCandidate[] }).commentCandidates = scan.candidates;
  }
}

function queueAddOrUpdateComment(
  context: OfficeContext,
  sheetName: string,
  address: string,
  text: string,
  cache: Map<string, OfficeComment | undefined>,
): void {
  const key = `${sheetName}!${normalizedCommentAddress(address)}`;
  if (!cache.has(key)) throw new Error("Comment scan cache unavailable");
  const cached = cache.get(key);
  if (cached !== undefined) {
    cached.content = text;
    cache.set(key, cached);
    return;
  }
  const targetSheet = sheet(context, sheetName);
  const comments = targetSheet.comments;
  if (comments === undefined) throw new Error("Office comments unavailable");
  const comment = comments.add(targetSheet.getRange(address), text);
  cache.set(key, comment);
}

type OfficePrecondition = {
  readonly kind: "object" | "range" | "chart_source" | "chart_collection" | "comments";
  readonly target: Record<string, unknown>;
  readonly expectedExists?: boolean;
  readonly expectedTitle?: string;
  readonly expectedRowCount?: number;
  readonly expectedColumnCount?: number;
  readonly expectedAddress?: string;
  readonly expectedAreaCount?: number;
  readonly expectedSentinel?: ScalarValue;
  readonly hasExpectedSentinel?: boolean;
  readonly expectedNumberFormat?: string;
  readonly expectedSourceFormulas?: readonly (readonly string[])[];
  readonly commentSheet?: string;
  readonly commentCandidates?: readonly OfficeCommentCandidate[];
};

function queueManifestPreconditions(
  context: OfficeContext,
  manifest: WorkbookPlanPreconditionManifest,
): readonly OfficePrecondition[] {
  if (manifest.version !== 1) throw new Error("Unsupported workbook precondition manifest");
  const preconditions: OfficePrecondition[] = [];
  const worksheets = context.workbook.worksheets;
  const tables = context.workbook.tables;
  if (worksheets.getItemOrNullObject === undefined || tables?.getItemOrNullObject === undefined) {
    throw new Error("Required Office existence readback APIs are unavailable");
  }

  for (const expected of manifest.sheets) {
    const target = worksheets.getItemOrNullObject(expected.name);
    if (target === undefined || typeof target.load !== "function") throw new Error("Worksheet precondition readback unavailable");
    target.load("isNullObject");
    preconditions.push({ kind: "object", target: target as Record<string, unknown>, expectedExists: expected.exists });
  }

  for (const expected of manifest.tables) {
    const target = tables.getItemOrNullObject(expected.name);
    if (target === undefined || typeof target.load !== "function") throw new Error("Table precondition readback unavailable");
    target.load("isNullObject");
    preconditions.push({ kind: "object", target: target as Record<string, unknown>, expectedExists: expected.exists });
    if (expected.exists) {
      const expectedRowCount = expected.rowCount;
      const expectedColumnCount = expected.columnCount;
      if (typeof expectedRowCount !== "number" || !Number.isSafeInteger(expectedRowCount) || expectedRowCount <= 0 ||
        typeof expectedColumnCount !== "number" || !Number.isSafeInteger(expectedColumnCount) || expectedColumnCount <= 0) {
        throw new Error("Table dimension precondition is incomplete");
      }
      const range = target.getRange?.();
      if (range === undefined || typeof range.load !== "function") throw new Error("Table dimension readback unavailable");
      range.load(["rowCount", "columnCount", "address"]);
      preconditions.push({
        kind: "range",
        target: range as Record<string, unknown>,
        expectedRowCount: expected.rowCount,
        expectedColumnCount: expected.columnCount,
        ...(expected.address === undefined ? {} : { expectedAddress: expected.address }),
      });
    }
  }

  for (const expected of manifest.ranges) {
    const targetSheet = worksheets.getItem(expected.sheet);
    if (targetSheet === undefined || typeof targetSheet.getRange !== "function") throw new Error("Range precondition readback unavailable");
    const target = targetSheet.getRange(expected.address);
    if (target === undefined || typeof target.load !== "function") throw new Error("Range precondition readback unavailable");
    if (!Number.isSafeInteger(expected.rowCount) || expected.rowCount <= 0 ||
      !Number.isSafeInteger(expected.columnCount) || expected.columnCount <= 0) {
      throw new Error("Range dimension precondition is incomplete");
    }
    const fields: string[] = ["rowCount", "columnCount", "address"];
    if (Object.prototype.hasOwnProperty.call(expected, "sentinel") || expected.numberFormat !== undefined) fields.push("values");
    if (expected.numberFormat !== undefined) fields.push("numberFormat");
    if (expected.sourceFormulas !== undefined) fields.push("formulas");
    target.load(fields);
    preconditions.push({
      kind: "range",
      target: target as Record<string, unknown>,
      expectedRowCount: expected.rowCount,
      expectedColumnCount: expected.columnCount,
      expectedAddress: expected.resolvedAddress ?? expected.address,
      ...(Object.prototype.hasOwnProperty.call(expected, "sentinel")
        ? { expectedSentinel: expected.sentinel, hasExpectedSentinel: true }
        : {}),
      ...(expected.numberFormat === undefined ? {} : { expectedNumberFormat: expected.numberFormat }),
      ...(expected.sourceFormulas === undefined ? {} : { expectedSourceFormulas: expected.sourceFormulas }),
    });
  }

  for (const expected of manifest.chartSources) {
    const targetSheet = worksheets.getItem(expected.sheet);
    if (targetSheet === undefined || typeof targetSheet.getRanges !== "function") throw new Error("Chart source readback unavailable");
    const target = targetSheet.getRanges(expected.address);
    if (target === undefined || typeof target.load !== "function") {
      throw new Error("Chart source contiguity readback unavailable");
    }
    target.load(["address", "areaCount"]);
    preconditions.push({
      kind: "chart_source",
      target: target as Record<string, unknown>,
      expectedAddress: expected.resolvedAddress ?? expected.address,
      expectedAreaCount: expected.areaCount,
    });
  }

  for (const expected of manifest.charts) {
    const targetSheet = worksheets.getItem(expected.sheet);
    const charts = targetSheet?.charts;
    if (charts === undefined || typeof charts.load !== "function") throw new Error("Chart existence readback unavailable");
    charts.load("items/title/text");
    preconditions.push({
      kind: "chart_collection",
      target: charts as Record<string, unknown>,
      expectedExists: expected.exists,
      expectedTitle: expected.title,
    });
  }

  for (const expected of manifest.comments) {
    const targetSheet = worksheets.getItem(expected.sheet);
    const comments = targetSheet?.comments;
    if (comments === undefined) throw new Error("Comment existence readback unavailable");
    preconditions.push({
      kind: "comments",
      target: comments as unknown as Record<string, unknown>,
      commentSheet: expected.sheet,
      expectedAddress: expected.address,
      expectedExists: expected.exists,
      ...(expected.text === undefined ? {} : { expectedSentinel: expected.text as ScalarValue, hasExpectedSentinel: true }),
    });
  }
  return preconditions;
}

function preconditionsFresh(preconditions: readonly OfficePrecondition[]): boolean {
  return preconditions.every((precondition) => {
    const target = precondition.target;
    if (precondition.kind === "comments") {
      const candidates = precondition.commentCandidates;
      if (candidates === undefined || precondition.expectedExists === undefined || precondition.expectedAddress === undefined) return false;
      const expectedAddress = normalizedCommentAddress(precondition.expectedAddress);
      const match = candidates.find(({ location }) => typeof location.address === "string" && normalizedCommentAddress(location.address) === expectedAddress);
      if ((match !== undefined) !== precondition.expectedExists) return false;
      return precondition.hasExpectedSentinel !== true || match?.comment.content === precondition.expectedSentinel;
    }
    if (precondition.kind === "object") {
      if (typeof target.isNullObject !== "boolean" || precondition.expectedExists === undefined) return false;
      if ((target.isNullObject !== true) !== precondition.expectedExists) return false;
      if (precondition.hasExpectedSentinel === true && target.content !== precondition.expectedSentinel) return false;
      return true;
    }
    if (precondition.kind === "chart_collection") {
      const items = target.items;
      if (!Array.isArray(items) || precondition.expectedExists === undefined || precondition.expectedTitle === undefined) return false;
      const exists = items.some((chart) => {
        if (chart === null || typeof chart !== "object") return false;
        const title = (chart as Record<string, unknown>).title;
        return title !== null && typeof title === "object" && (title as Record<string, unknown>).text === precondition.expectedTitle;
      });
      return exists === precondition.expectedExists;
    }
    if (precondition.kind === "chart_source") {
      if (precondition.expectedAddress !== undefined && target.address !== precondition.expectedAddress) return false;
      return typeof target.areaCount === "number" && target.areaCount === precondition.expectedAreaCount;
    }
    if (typeof target.rowCount !== "number" || typeof target.columnCount !== "number") return false;
    if (precondition.expectedRowCount !== undefined && target.rowCount !== precondition.expectedRowCount) return false;
    if (precondition.expectedColumnCount !== undefined && target.columnCount !== precondition.expectedColumnCount) return false;
    if (precondition.expectedAddress !== undefined && target.address !== precondition.expectedAddress) return false;
    if (precondition.hasExpectedSentinel === true) {
      const values = target.values;
      if (!Array.isArray(values) || !Array.isArray(values[0]) || !Object.is(values[0][0], precondition.expectedSentinel)) return false;
    }
    if (precondition.expectedNumberFormat !== undefined) {
      const formats = target.numberFormat;
      if (!Array.isArray(formats) || formats.some((row) => !Array.isArray(row) || row.some((value) => value !== precondition.expectedNumberFormat))) return false;
    }
    if (precondition.expectedSourceFormulas !== undefined && !jsonEqual(target.formulas, precondition.expectedSourceFormulas)) return false;
    return true;
  });
}

function queueFormatReadback(range: OfficeRange, format: Record<string, unknown>): void {
  if (typeof range.load !== "function" || range.format === undefined) throw new Error("Office format readback unavailable");
  range.load(["values", "formulas", "numberFormat"]);
  const owner = range.format;
  const ownerFields: string[] = [];
  for (const key of ["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation"] as const) {
    if (format[key] !== undefined) ownerFields.push(key);
  }
  if (format.columnAutofit === true) ownerFields.push("columnWidth");
  if (format.rowAutofit === true) ownerFields.push("rowHeight");
  if (ownerFields.length > 0) {
    if (typeof owner.load !== "function") throw new Error("Office range-format readback unavailable");
    owner.load(ownerFields);
  }
  const fontFields = [
    ["bold", "bold"], ["italic", "italic"], ["fontColor", "color"], ["fontSize", "size"],
    ["fontName", "name"], ["underline", "underline"],
  ] as const;
  const font = owner.font;
  if (fontFields.some(([key]) => format[key] !== undefined)) {
    if (font === undefined || typeof font.load !== "function") throw new Error("Office font readback unavailable");
    font.load(fontFields.filter(([key]) => format[key] !== undefined).map(([, field]) => field));
  }
  if (format.fillColor !== undefined) {
    const fill = owner.fill;
    if (fill === undefined || typeof fill.load !== "function") throw new Error("Office fill readback unavailable");
    fill.load("color");
  }
  if (format.borders === true) {
    const borders = owner.borders;
    if (borders === undefined || typeof borders.load !== "function") throw new Error("Office border readback unavailable");
    borders.load(["items/sideIndex", "items/style", "items/color", "items/weight"]);
  }
}

function queueReconciliationRead(context: OfficeContext, operation: WorkbookPlanEffectOperation): OfficeReadRequest | undefined {
  const args = operation.args;
  switch (operation.kind) {
    case "write_values":
    case "write_formula":
    case "fill_formula":
    case "format_range": {
      const address = operation.kind === "fill_formula" ? args.targetAddress : args.address;
      if (typeof args.sheet !== "string" || typeof address !== "string") return undefined;
      const range = sheet(context, args.sheet).getRange(address);
      if (operation.kind === "format_range") {
        queueFormatReadback(range, recordArg(args, "format"));
      } else {
        requireLoad(range, ["values", "formulas", "numberFormat"]);
      }
      return { kind: operation.kind, target: range };
    }
    case "create_table": {
      const tables = context.workbook.tables;
      const objectName = internalObjectName("Table", operation.operationId);
      if (tables === undefined || objectName === undefined || typeof tables.getItemOrNullObject !== "function") return undefined;
      const target = tables.getItemOrNullObject(objectName);
      requireLoad(target, ["isNullObject", "name", "showHeaders", "showTotals"]);
      const range = target.getRange?.();
      if (range === undefined) return undefined;
      requireLoad(range, ["address", "rowCount", "columnCount"]);
      return { kind: "table_created", target, range };
    }
    case "sort_table": {
      const tables = context.workbook.tables;
      if (tables === undefined) return undefined;
      const target = tables.getItem(stringArg(args, "tableName"));
      if (target.sort === undefined || typeof target.sort.load !== "function") return undefined;
      target.sort.load("fields");
      const range = target.getDataBodyRange?.();
      if (range === undefined) return undefined;
      requireLoad(range, "values");
      return { kind: "table_sort", target: target.sort, range };
    }
    case "filter_table":
    case "reset_table_filter": {
      const tables = context.workbook.tables;
      if (tables === undefined) return undefined;
      const target = tables.getItem(stringArg(args, "tableName"));
      const filter = target.columns?.getItemAt(numberArg(args, "columnIndex")).filter;
      if (filter === undefined || typeof filter.load !== "function") return undefined;
      filter.load("criteria");
      return { kind: "table_filter", target: filter };
    }
    case "set_validation_list": {
      const range = sheet(context, stringArg(args, "sheet")).getRange(stringArg(args, "address"));
      const validation = range.dataValidation as Record<string, unknown> | undefined;
      if (validation === undefined || typeof validation.load !== "function") return undefined;
      validation.load(["type", "rule", "prompt"]);
      return { kind: "validation", target: validation };
    }
    case "conditional_format": {
      const range = sheet(context, stringArg(args, "sheet")).getRange(stringArg(args, "address"));
      const formats = range.conditionalFormats;
      if (formats === undefined || typeof formats.load !== "function") return undefined;
      const fields = stringArg(args, "ruleType") === "cellValue"
        ? ["items/id", "items/type", "items/cellValue/rule", "items/cellValue/format/fill/color", "items/cellValue/format/font/color"]
        : ["items/id", "items/type", "items/textComparison/rule", "items/textComparison/format/fill/color", "items/textComparison/format/font/color"];
      formats.load(fields);
      return { kind: "conditional_format", target: formats };
    }
    case "freeze_panes": {
      const panes = sheet(context, stringArg(args, "sheet")).freezePanes;
      if (panes === undefined || typeof panes.getLocationOrNullObject !== "function") return undefined;
      const target = panes.getLocationOrNullObject();
      requireLoad(target, ["isNullObject", "address", "rowCount", "columnCount"]);
      return { kind: "freeze_panes", target };
    }
    case "create_chart": {
      const charts = sheet(context, stringArg(args, "sheet")).charts;
      const objectName = internalObjectName("Chart", operation.operationId);
      if (charts === undefined || objectName === undefined) return undefined;
      const target = charts.getItemOrNullObject?.(objectName);
      if (target === undefined) return undefined;
      requireLoad(target, ["isNullObject", "name", "chartType", "title/text", "left", "top", "width", "height"]);
      return { kind: "chart", target };
    }
    case "create_sheet":
    case "copy_sheet": {
      const target = context.workbook.worksheets.getItemOrNullObject?.(
        stringArg(args, operation.kind === "create_sheet" ? "name" : "newName"),
      );
      if (target === undefined) return undefined;
      requireLoad(target, "isNullObject");
      return { kind: "sheet_absent", target };
    }
    case "rename_sheet": {
      const newTarget = context.workbook.worksheets.getItemOrNullObject?.(stringArg(args, "newName"));
      const oldTarget = context.workbook.worksheets.getItemOrNullObject?.(stringArg(args, "oldName"));
      if (newTarget === undefined || oldTarget === undefined) return undefined;
      requireLoad(newTarget, "isNullObject");
      requireLoad(oldTarget, "isNullObject");
      return { kind: "sheet_renamed", target: newTarget, secondary: oldTarget };
    }
    case "add_comment": {
      const targetSheet = sheet(context, stringArg(args, "sheet"));
      const comments = targetSheet.comments;
      if (comments === undefined) return undefined;
      return {
        kind: "comments",
        target: comments,
        commentSheet: stringArg(args, "sheet"),
        expectedAddress: stringArg(args, "address"),
      };
    }
    case "calculate": {
      const application = context.workbook.application;
      if (application === undefined) return undefined;
      requireLoad(application, "calculationState");
      return { kind: "calculation", target: application };
    }
  }
}

function conditionalReadbackMatches(
  operation: WorkbookPlanEffectOperation,
  target: Record<string, unknown>,
  identityBefore?: OperationIdentityBefore,
): boolean {
  const items = target.items;
  if (!Array.isArray(items) || items.length === 0 || items.length > 512) return false;
  const identity = identityBefore?.conditionalFormatIds;
  if (identity === undefined) return false;
  const candidateItems: Record<string, unknown>[] = [];
  for (const value of items) {
    if (value === null || typeof value !== "object") return false;
    const record = value as Record<string, unknown>;
    if (typeof record.id !== "string") return false;
    if (!identity.has(record.id)) candidateItems.push(record);
  }
  if (candidateItems.length !== 1) return false;
  const item = candidateItems[0]!;
  const args = operation.args;
  const ruleType = stringArg(args, "ruleType");
  const expectedType = normalizeConditionalType(ruleType);
  if (item.type !== expectedType) return false;
  const typed = ruleType === "cellValue" || ruleType === "CellValue" ? item.cellValue : item.textComparison;
  if (typed === null || typeof typed !== "object") return false;
  const typedRecord = typed as Record<string, unknown>;
  const actualRule = typedRecord.rule;
  if (actualRule === null || typeof actualRule !== "object" || Array.isArray(actualRule)) return false;
  const expectedRule = ruleType === "cellValue" || ruleType === "CellValue"
    ? { operator: normalizeCellValueOperator(typeof args.operator === "string" ? args.operator : "GreaterThan"), formula1: stringArg(args, "formula1"), ...(typeof args.formula2 === "string" ? { formula2: args.formula2 } : {}) }
    : { operator: normalizeTextComparisonOperator(typeof args.operator === "string" ? args.operator : "Contains"), text: stringArg(args, "text") };
  if (!jsonEqual(actualRule, expectedRule)) return false;
  const format = typedRecord.format;
  if (format === null || typeof format !== "object") return false;
  const formatRecord = format as Record<string, unknown>;
  const fill = formatRecord.fill as Record<string, unknown> | undefined;
  const font = formatRecord.font as Record<string, unknown> | undefined;
  return (args.fillColor === undefined || fill?.color === args.fillColor) &&
    (args.fontColor === undefined || font?.color === args.fontColor);
}

function isExactSortField(value: unknown): value is { readonly key: number; readonly ascending: boolean } {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  return Object.keys(record).sort().join(",") === "ascending,key" &&
    typeof record.key === "number" && Number.isSafeInteger(record.key) && record.key >= 0 &&
    typeof record.ascending === "boolean";
}

function isSortScalar(value: unknown): value is ScalarValue {
  return value === null || typeof value === "boolean" ||
    typeof value === "string" && value.length <= 32_767 ||
    typeof value === "number" && Number.isFinite(value);
}

function cloneValueMatrix(matrix: ValueMatrix): ValueMatrix {
  return matrix.map((row) => [...row]);
}

function isValidSortBody(value: unknown, columnIndex: unknown): ValueMatrix | undefined {
  if (!Array.isArray(value) || value.length < 1 || value.length > MAX_SORT_BODY_ROWS ||
    typeof columnIndex !== "number" || !Number.isSafeInteger(columnIndex) || columnIndex < 0) return undefined;
  const rows: ScalarValue[][] = [];
  let columnCount: number | undefined;
  const keys = new Set<number>();
  let cells = 0;
  for (const candidate of value) {
    if (!Array.isArray(candidate) || candidate.length < 1 || candidate.length > MAX_SORT_BODY_COLUMNS ||
      columnCount !== undefined && candidate.length !== columnCount) return undefined;
    columnCount = candidate.length;
    cells += candidate.length;
    if (cells > MAX_SORT_BODY_CELLS || columnIndex >= candidate.length) return undefined;
    const row: ScalarValue[] = [];
    for (const cell of candidate) {
      if (!isSortScalar(cell)) return undefined;
      row.push(cell);
    }
    const key = row[columnIndex];
    if (typeof key !== "number" || !Number.isFinite(key) || keys.has(key)) return undefined;
    keys.add(key);
    rows.push(row);
  }
  return rows;
}

function expectedSortBodyFor(body: ValueMatrix, columnIndex: unknown, ascending: boolean): ValueMatrix | undefined {
  const valid = isValidSortBody(body, columnIndex);
  if (valid === undefined || typeof columnIndex !== "number") return undefined;
  return [...valid]
    .sort((left, right) => {
      const leftKey = left[columnIndex];
      const rightKey = right[columnIndex];
      if (typeof leftKey !== "number" || typeof rightKey !== "number") return 0;
      const comparison = leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
      return ascending ? comparison : -comparison;
    })
    .map((row) => [...row]);
}

function validationReadbackMatches(
  operation: WorkbookPlanEffectOperation,
  target: Record<string, unknown>,
): boolean {
  if (target.type !== "List") return false;
  const rule = target.rule;
  if (rule === null || typeof rule !== "object" || Array.isArray(rule)) return false;
  const ruleRecord = rule as Record<string, unknown>;
  if (Object.keys(ruleRecord).length !== 1 || ruleRecord.list === null || typeof ruleRecord.list !== "object" || Array.isArray(ruleRecord.list)) return false;
  const list = ruleRecord.list as Record<string, unknown>;
  if (Object.keys(list).sort().join(",") !== "inCellDropDown,source" || list.inCellDropDown !== true || list.source !== operation.args.source) return false;
  if (operation.args.promptTitle === undefined && operation.args.promptMessage === undefined) return true;
  const prompt = target.prompt;
  if (prompt === null || typeof prompt !== "object" || Array.isArray(prompt)) return false;
  const promptRecord = prompt as Record<string, unknown>;
  return Object.keys(promptRecord).sort().join(",") === "message,showPrompt,title" &&
    promptRecord.showPrompt === true &&
    promptRecord.title === (operation.args.promptTitle ?? "") &&
    promptRecord.message === (operation.args.promptMessage ?? "");
}

const RANGE_BORDER_SIDES = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const;

type FormatPath = readonly ["font" | "fill" | "range", string];

function formatReadbackMatches(
  operation: WorkbookPlanEffectOperation,
  target: Record<string, unknown>,
): boolean {
  const format = recordArg(operation.args, "format");
  const actual = target.format as Record<string, unknown> | undefined;
  if (actual === undefined) return false;
  const requested: Readonly<Record<string, FormatPath>> = {
    bold: ["font", "bold"],
    italic: ["font", "italic"],
    fontColor: ["font", "color"],
    fontSize: ["font", "size"],
    fontName: ["font", "name"],
    underline: ["font", "underline"],
    fillColor: ["fill", "color"],
    horizontalAlignment: ["range", "horizontalAlignment"],
    verticalAlignment: ["range", "verticalAlignment"],
    wrapText: ["range", "wrapText"],
    indentLevel: ["range", "indentLevel"],
    textOrientation: ["range", "textOrientation"],
  };
  for (const [key, [owner, property]] of Object.entries(requested)) {
    if (format[key] === undefined) continue;
    const source = owner === "range" ? actual : actual[owner] as Record<string, unknown> | undefined;
    if (source === undefined || !Object.is(source[property], format[key])) return false;
  }
  if (typeof format.numberFormat === "string" && !jsonEqual(target.numberFormat, dimensionedNumberFormat(stringArg(operation.args, "address"), format.numberFormat))) return false;
  if (format.columnAutofit === true) {
    const width = actual.columnWidth;
    // Excel exposes the resulting width, not an "autofit enabled" flag. A
    // finite width is the bounded host-observable postcondition we can prove.
    if (typeof width !== "number" || !Number.isFinite(width) || width <= 0) return false;
  }
  if (format.rowAutofit === true) {
    const height = actual.rowHeight;
    // Excel exposes the resulting height, not an "autofit enabled" flag.
    if (typeof height !== "number" || !Number.isFinite(height) || height <= 0) return false;
  }
  if (format.borders === true) {
    const borders = actual.borders as Record<string, unknown> | undefined;
    const items = borders?.items;
    if (!Array.isArray(items) || items.length !== RANGE_BORDER_SIDES.length) return false;
    const bySide = new Map<string, Record<string, unknown>>();
    for (const item of items) {
      if (item === null || typeof item !== "object") return false;
      const record = item as Record<string, unknown>;
      if (typeof record.sideIndex !== "string" || bySide.has(record.sideIndex)) return false;
      bySide.set(record.sideIndex, record);
    }
    if (RANGE_BORDER_SIDES.some((side) => !bySide.has(side))) return false;
    for (const side of RANGE_BORDER_SIDES) {
      const border = bySide.get(side)!;
      if (border.style !== "Continuous") return false;
      if (format.borderColor !== undefined && border.color !== format.borderColor) return false;
      if (format.borderWeight !== undefined && border.weight !== format.borderWeight) return false;
    }
  }
  return true;
}

function reconcileOperation(
  operation: WorkbookPlanEffectOperation,
  request: OfficeReadRequest | undefined,
): "committed" | "not_committed" | "unknown_within_batch" {
  if (request === undefined) return "unknown_within_batch";
  const target = request.target as Record<string, unknown> | undefined;
  if (target === undefined) return "unknown_within_batch";
  const args = operation.args;
  if (operation.kind === "write_values") {
    if (jsonEqual(target.values, args.values)) return "committed";
    const preValues = request.identityBefore?.preValues;
    return preValues !== undefined && jsonEqual(target.values, preValues) ? "not_committed" : "unknown_within_batch";
  }
  if (operation.kind === "write_formula") return jsonEqual(target.formulas, args.values) ? "committed" : "unknown_within_batch";
  if (operation.kind === "fill_formula") {
    const expected = request.expectedFormulas;
    return expected !== undefined && jsonEqual(target.formulas, expected) ? "committed" : "unknown_within_batch";
  }
  if (operation.kind === "format_range") return formatReadbackMatches(operation, target) ? "committed" : "unknown_within_batch";
  if (operation.kind === "set_validation_list" && request.kind === "validation") return validationReadbackMatches(operation, target) ? "committed" : "unknown_within_batch";
  if (operation.kind === "conditional_format" && request.kind === "conditional_format") return conditionalReadbackMatches(operation, target, request.identityBefore) ? "committed" : "unknown_within_batch";
  if (operation.kind === "create_table" && request.kind === "table_created") {
    const table = target;
    const range = request.range as Record<string, unknown> | undefined;
    const name = internalObjectName("Table", operation.operationId);
    if (table.isNullObject === true) return "not_committed";
    if (request.identityBefore?.objectAbsent !== true) return "unknown_within_batch";
    if (name === undefined || table.name !== name || table.showHeaders !== (args.hasHeaders === true) || table.showTotals !== false || range === undefined) return "unknown_within_batch";
    if (typeof range.rowCount !== "number" || typeof range.columnCount !== "number" || range.rowCount < 1 || range.columnCount < 1) return "unknown_within_batch";
    return addressTailEquals(range.address, stringArg(args, "address")) ? "committed" : "unknown_within_batch";
  }
  if (operation.kind === "sort_table" && request.kind === "table_sort") {
    const fields = target.fields;
    if (!Array.isArray(fields) || fields.length > 1) return "unknown_within_batch";
    if (fields.length === 0) return "not_committed";
    const field = fields[0];
    if (!isExactSortField(field) || field.key !== args.columnIndex || field.ascending !== (args.ascending !== false)) return "unknown_within_batch";
    const identity = request.identityBefore;
    const expectedBody = identity?.expectedSortBody;
    const actualBody = request.range?.values;
    if (expectedBody === undefined || !isValidSortBody(actualBody, args.columnIndex)) return "unknown_within_batch";
    if (jsonEqual(actualBody, expectedBody)) return "committed";
    return identity?.preSortBody !== undefined && jsonEqual(actualBody, identity.preSortBody)
      ? "not_committed"
      : "unknown_within_batch";
  }
  if ((operation.kind === "filter_table" || operation.kind === "reset_table_filter") && request.kind === "table_filter") {
    return filterReadbackMatches(operation, target.criteria) ? "committed" : "unknown_within_batch";
  }
  if (operation.kind === "freeze_panes" && request.kind === "freeze_panes") return freezeReadbackMatches(operation, target) ? "committed" : "unknown_within_batch";
  if (request.kind === "comments") {
    const candidates = request.commentCandidates;
    if (candidates === undefined || request.expectedAddress === undefined) return "unknown_within_batch";
    const expectedAddress = normalizedCommentAddress(request.expectedAddress);
    const match = candidates.find(({ location }) => typeof location.address === "string" && normalizedCommentAddress(location.address) === expectedAddress);
    if (match !== undefined && match.comment.content === args.text) return "committed";
    const preComment = request.identityBefore?.preComment;
    if (preComment === undefined) return "unknown_within_batch";
    if (!preComment.exists && match === undefined) return "not_committed";
    if (preComment.exists && match !== undefined && typeof preComment.text === "string" && match.comment.content === preComment.text) return "not_committed";
    return "unknown_within_batch";
  }
  if (request.kind === "chart" && operation.kind === "create_chart") {
    const name = internalObjectName("Chart", operation.operationId);
    if (target.isNullObject === true) return "not_committed";
    if (request.identityBefore?.objectAbsent !== true) return "unknown_within_batch";
    const title = target.title as Record<string, unknown> | undefined;
    if (name === undefined || target.name !== name || target.chartType !== args.chartType || title?.text !== args.title) return "unknown_within_batch";
    if (request.identityBefore.sourceAreaCount !== 1 || request.identityBefore.sourceAddress !== normalizedAddress(stringArg(args, "dataRange"))) return "unknown_within_batch";
    if (typeof target.left !== "number" || !Number.isFinite(target.left) || target.left < 0 ||
      typeof target.top !== "number" || !Number.isFinite(target.top) || target.top < 0 ||
      typeof target.width !== "number" || !Number.isFinite(target.width) || target.width <= 0 ||
      typeof target.height !== "number" || !Number.isFinite(target.height) || target.height <= 0) return "unknown_within_batch";
    return "committed";
  }
  if (request.kind === "table" || request.kind === "chart" || request.kind === "sheet_absent") {
    return target.isNullObject === false ? "committed" : target.isNullObject === true ? "not_committed" : "unknown_within_batch";
  }
  if (request.kind === "sheet_renamed") {
    const oldTarget = request.secondary as Record<string, unknown> | undefined;
    if (oldTarget === undefined || typeof oldTarget.isNullObject !== "boolean" || typeof target.isNullObject !== "boolean") return "unknown_within_batch";
    if (target.isNullObject === false && oldTarget.isNullObject === true) return "committed";
    if (target.isNullObject === true && oldTarget.isNullObject === false) return "not_committed";
    return "unknown_within_batch";
  }
  if (request.kind === "sheet_present") return target.isNullObject === false ? "committed" : "not_committed";
  if (request.kind === "calculation") return target.calculationState === "Done" ? "committed" : "unknown_within_batch";
  return "unknown_within_batch";
}

function addressTailEquals(actual: unknown, expected: string): boolean {
  if (typeof actual !== "string") return false;
  return normalizedAddress(actual) === normalizedAddress(expected);
}

function normalizedAddress(value: string): string {
  return (value.split("!").at(-1) ?? value).replace(/\$/g, "").toUpperCase();
}

function filterReadbackMatches(operation: WorkbookPlanEffectOperation, criteriaValue: unknown): boolean {
  if (criteriaValue === null || typeof criteriaValue !== "object" || Array.isArray(criteriaValue)) return false;
  const criteria = criteriaValue as Record<string, unknown>;
  const keys = Object.keys(criteria).sort();
  if (operation.kind === "reset_table_filter") {
    return keys.length === 1 && keys[0] === "filterOn" && criteria.filterOn === "None";
  }
  if (operation.args.values !== undefined) {
    const expectedValues = operation.args.values;
    if (!Array.isArray(expectedValues) || expectedValues.some((value) => typeof value !== "string") ||
      keys.join(",") !== "filterOn,values" || criteria.filterOn !== "Values" || !Array.isArray(criteria.values)) return false;
    return criteria.values.length === expectedValues.length &&
      criteria.values.every((value, index) => typeof value === "string" && value === expectedValues[index]);
  }
  if (keys.some((key) => !["filterOn", "criterion1", "criterion2", "operator"].includes(key)) ||
    criteria.filterOn !== "Custom" || typeof operation.args.criterion1 !== "string" || criteria.criterion1 !== operation.args.criterion1) return false;
  const hasExpectedCriterion2 = Object.prototype.hasOwnProperty.call(operation.args, "criterion2");
  if (hasExpectedCriterion2 !== Object.prototype.hasOwnProperty.call(criteria, "criterion2")) return false;
  if (hasExpectedCriterion2 && (typeof operation.args.criterion2 !== "string" || criteria.criterion2 !== operation.args.criterion2)) return false;
  const actualOperator = criteria.operator === undefined ? "And" : criteria.operator;
  const expectedOperator = operation.args.operator === undefined ? "And" : operation.args.operator;
  return (actualOperator === "And" || actualOperator === "Or") && actualOperator === expectedOperator;
}

function freezeReadbackMatches(operation: WorkbookPlanEffectOperation, target: Record<string, unknown>): boolean {
  const rows = numberArg(operation.args, "rows");
  const columns = numberArg(operation.args, "columns");
  if (rows === 0 && columns === 0) return target.isNullObject === true;
  if (target.isNullObject !== false) return false;
  if (rows > 0 && target.rowCount !== rows) return false;
  if (columns > 0 && target.columnCount !== columns) return false;
  return true;
}

function jsonEqual(left: unknown, right: unknown): boolean {
  try { return JSON.stringify(left) === JSON.stringify(right); } catch { return false; }
}

type OfficeReadRequest = {
  readonly kind: string;
  readonly target?: unknown;
  readonly secondary?: unknown;
  readonly range?: OfficeRange;
  readonly expectedFormulas?: unknown;
  readonly identityBefore?: OperationIdentityBefore;
  readonly commentSheet?: string;
  readonly expectedAddress?: string;
  readonly commentCandidates?: readonly OfficeCommentCandidate[];
};

type OperationIdentityBefore = {
  readonly operation?: WorkbookPlanEffectOperation;
  readonly objectAbsent?: boolean;
  readonly conditionalFormatIds?: ReadonlySet<string>;
  readonly preValues?: ValueMatrix;
  readonly preComment?: { readonly exists: boolean; readonly text?: string };
  readonly sourceRange?: OfficeRange;
  readonly sourceAddress?: string;
  readonly sourceAreaCount?: number;
  readonly preSortBody?: ValueMatrix;
  readonly expectedSortBody?: ValueMatrix;
};

async function captureOperationIdentities(
  context: OfficeContext,
  operations: readonly WorkbookPlanEffectOperation[],
  identities: Map<string, OperationIdentityBefore>,
  commentScans: OfficeCommentScans,
): Promise<void> {
  const pending: Array<{
    readonly operation: WorkbookPlanEffectOperation;
    readonly target: Record<string, unknown>;
    readonly collection?: Record<string, unknown>;
    readonly capturesValues?: boolean;
    readonly capturesSortBody?: boolean;
    readonly sourceRange?: OfficeRange;
    readonly sourceAreas?: OfficeRangeAreas;
  }> = [];
  for (const operation of operations) {
    if (operation.kind === "write_values") {
      const range = sheet(context, stringArg(operation.args, "sheet")).getRange(stringArg(operation.args, "address"));
      if (typeof range.load !== "function") throw new Error("Write-value pre-state readback unavailable");
      range.load("values");
      pending.push({ operation, target: range as Record<string, unknown>, capturesValues: true });
    } else if (operation.kind === "sort_table") {
      const columnIndex = operation.args.columnIndex;
      if (typeof columnIndex !== "number" || !Number.isSafeInteger(columnIndex) || columnIndex < 0) {
        throw new Error("Sort column index is invalid");
      }
      const tables = context.workbook.tables;
      if (tables === undefined) throw new Error("Table collection unavailable");
      const table = tables.getItem(stringArg(operation.args, "tableName"));
      const range = table.getDataBodyRange?.();
      if (range === undefined) throw new Error("Table sort data-body readback unavailable");
      requireLoad(range, "values");
      pending.push({ operation, target: range as Record<string, unknown>, capturesSortBody: true });
    } else if (operation.kind === "add_comment") {
      const sheetName = stringArg(operation.args, "sheet");
      const scan = commentScans.get(sheetName);
      if (scan === undefined) throw new Error("Comment pre-state scan unavailable");
      const comment = scan.index.get(normalizedCommentAddress(stringArg(operation.args, "address")));
      identities.set(operation.operationId, {
        operation,
        preComment: {
          exists: comment !== undefined,
          ...(comment === undefined ? {} : { text: comment.content }),
        },
      });
    }
    if (operation.kind === "create_table" || operation.kind === "create_chart") {
      const objectName = internalObjectName(operation.kind === "create_table" ? "Table" : "Chart", operation.operationId);
      if (objectName === undefined) throw new Error("Invalid trusted created-object identity");
      const targetSheet = operation.kind === "create_chart" ? sheet(context, stringArg(operation.args, "sheet")) : undefined;
      const target = operation.kind === "create_table"
        ? context.workbook.tables?.getItemOrNullObject?.(objectName)
        : targetSheet?.charts?.getItemOrNullObject?.(objectName);
      if (target === undefined) throw new Error("Created-object identity readback unavailable");
      const load = (target as Record<string, unknown>).load;
      if (typeof load !== "function") throw new Error("Created-object identity readback unavailable");
      load.call(target, "isNullObject");
      if (operation.kind === "create_chart") {
        const dataRange = stringArg(operation.args, "dataRange");
        const sourceRange = targetSheet?.getRange(dataRange);
        const sourceAreas = targetSheet?.getRanges?.(dataRange);
        if (sourceRange === undefined || sourceAreas === undefined) throw new Error("Chart source identity readback unavailable");
        requireLoad(sourceRange, "address");
        requireLoad(sourceAreas, ["address", "areaCount"]);
        pending.push({ operation, target: target as Record<string, unknown>, sourceRange, sourceAreas });
      } else {
        pending.push({ operation, target: target as Record<string, unknown> });
      }
    } else if (operation.kind === "conditional_format") {
      const range = sheet(context, stringArg(operation.args, "sheet")).getRange(stringArg(operation.args, "address"));
      const formats = range.conditionalFormats as unknown as Record<string, unknown> | undefined;
      if (formats === undefined || typeof formats.load !== "function") throw new Error("Conditional-format identity readback unavailable");
      formats.load("items/id");
      pending.push({ operation, target: formats, collection: formats });
    }
  }
  if (pending.length === 0) return;
  await context.sync();
  for (const item of pending) {
    if (item.capturesValues === true) {
      const values = item.target.values;
      if (!isScalarValueMatrix(values)) throw new Error("Write-value pre-state is malformed");
      identities.set(item.operation.operationId, { operation: item.operation, preValues: values.map((row) => [...row]) });
    } else if (item.capturesSortBody === true) {
      const columnIndex = item.operation.args.columnIndex;
      const values = item.target.values;
      const sortBody = isValidSortBody(values, columnIndex);
      if (sortBody === undefined) throw new Error("Sort data-body pre-state is not exactly comparable");
      const expectedSortBody = expectedSortBodyFor(sortBody, columnIndex, item.operation.args.ascending !== false);
      if (expectedSortBody === undefined) throw new Error("Sort data-body pre-state is not exactly sortable");
      identities.set(item.operation.operationId, {
        operation: item.operation,
        preSortBody: cloneValueMatrix(sortBody),
        expectedSortBody,
      });
    } else if (item.collection !== undefined) {
      const values = item.collection.items;
      if (!Array.isArray(values) || values.length > 512) throw new Error("Conditional-format collection exceeds conservative bound");
      const ids = new Set<string>();
      for (const value of values) {
        const id = value && typeof value === "object" ? (value as Record<string, unknown>).id : undefined;
        if (typeof id !== "string" || ids.has(id)) throw new Error("Conditional-format identity is malformed or ambiguous");
        ids.add(id);
      }
      identities.set(item.operation.operationId, { operation: item.operation, conditionalFormatIds: ids });
    } else {
      if (item.target.isNullObject !== true) throw new Error("Created-object identity collision");
      if (item.operation.kind === "create_chart") {
        if (item.sourceRange === undefined || item.sourceAreas === undefined ||
          typeof item.sourceRange.address !== "string" || typeof item.sourceAreas.address !== "string" || item.sourceAreas.areaCount !== 1 ||
          !addressTailEquals(item.sourceRange.address, stringArg(item.operation.args, "dataRange")) ||
          !addressTailEquals(item.sourceAreas.address, stringArg(item.operation.args, "dataRange")) ||
          normalizedAddress(item.sourceRange.address) !== normalizedAddress(item.sourceAreas.address)) {
          throw new Error("Chart source identity is malformed or non-contiguous");
        }
        identities.set(item.operation.operationId, {
          operation: item.operation,
          objectAbsent: true,
          sourceRange: item.sourceRange,
          sourceAddress: normalizedAddress(item.sourceRange.address),
          sourceAreaCount: item.sourceAreas.areaCount,
        });
      } else {
        identities.set(item.operation.operationId, { operation: item.operation, objectAbsent: true });
      }
    }
  }
}

function queueRead(context: OfficeContext, rule: VerificationRule): OfficeReadRequest | undefined {
  switch (rule.kind) {
    case "range_values_equal":
    case "range_formulas_equal":
    case "range_formula_pattern":
    case "range_number_format_equal":
    case "sentinel_unchanged": {
      const range = sheet(context, rule.sheet).getRange(rule.address);
      requireLoad(range, ["values", "formulas", "numberFormat"]);
      return { kind: "range", target: range };
    }
    case "table_exists": {
      const tables = context.workbook.tables;
      if (tables === undefined) return undefined;
      const target = tables.getItemOrNullObject?.(rule.tableName) ?? tables.getItem(rule.tableName);
      requireLoad(target, "isNullObject");
      return { kind: "table", target };
    }
    case "table_row_order": {
      const tables = context.workbook.tables;
      if (tables === undefined) return undefined;
      const target = tables.getItem(rule.tableName);
      const range = target.getDataBodyRange?.();
      if (range === undefined) return undefined;
      requireLoad(range, "values");
      return { kind: "table_row_order", target, range };
    }
    case "table_filter_state": {
      const tables = context.workbook.tables;
      if (tables === undefined) return undefined;
      const target = tables.getItem(rule.tableName);
      const filter = target.columns?.getItemAt(rule.columnIndex).filter;
      if (filter === undefined) return undefined;
      requireLoad(filter, "criteria");
      return { kind: "table_filter_state", target: filter };
    }
    case "sheet_exists": {
      const target = context.workbook.worksheets.getItemOrNullObject?.(rule.name);
      if (target === undefined) return undefined;
      requireLoad(target, "isNullObject");
      return { kind: "sheet", target };
    }
    case "chart_exists": {
      const charts = sheet(context, rule.sheet).charts;
      if (charts === undefined || typeof charts.load !== "function") return undefined;
      charts.load("items/title/text");
      return { kind: "chart_collection", target: charts };
    }
    case "comment_exists": {
      const comments = sheet(context, rule.sheet).comments;
      if (comments === undefined) return undefined;
      return { kind: "comments", target: comments, commentSheet: rule.sheet, expectedAddress: rule.address };
    }
    case "calculation_completed": {
      const application = context.workbook.application;
      if (application === undefined) return undefined;
      requireLoad(application, "calculationState");
      return { kind: "calculation", target: application };
    }
  }
}

function readSnapshot(rule: VerificationRule, request: OfficeReadRequest | undefined): unknown {
  if (request === undefined) return undefined;
  const value = (request.target ?? {}) as Record<string, unknown>;
  switch (rule.kind) {
    case "range_values_equal": return { kind: "range_values", values: value.values };
    case "range_formulas_equal":
    case "range_formula_pattern": return { kind: "range_formulas", formulas: value.formulas };
    case "range_number_format_equal": return { kind: "range_number_format", numberFormat: value.numberFormat };
    case "sentinel_unchanged": return { kind: "sentinel", value: Array.isArray(value.values) && Array.isArray(value.values[0]) ? value.values[0][0] : null };
    case "table_exists": return { kind: "table_exists", exists: value.isNullObject !== true };
    case "table_row_order": {
      const rows = (request.range as Record<string, unknown> | undefined)?.values;
      if (!Array.isArray(rows)) return undefined;
      return { kind: "table_row_order", order: rows.map((row) => Array.isArray(row) ? String(row[0] ?? "") : "") };
    }
    case "table_filter_state": {
      const criteria = value.criteria;
      if (criteria === null || typeof criteria !== "object" || Array.isArray(criteria)) return undefined;
      const record = criteria as Record<string, unknown>;
      const keys = Object.keys(record).sort();
      if (keys.length === 1 && keys[0] === "filterOn" && record.filterOn === "None") {
        return { kind: "table_filter_state", state: { columnIndex: rule.columnIndex, state: "unfiltered" } };
      }
      if (record.filterOn === "Values" && keys.join(",") === "filterOn,values" && Array.isArray(record.values) && record.values.every((item) => typeof item === "string")) {
        return { kind: "table_filter_state", state: { columnIndex: rule.columnIndex, state: "filtered", values: record.values as string[] } };
      }
      if (record.filterOn !== "Custom" || keys.some((key) => !["filterOn", "criterion1", "criterion2", "operator"].includes(key)) ||
        !Object.prototype.hasOwnProperty.call(record, "criterion1") ||
        typeof record.criterion1 !== "string" || record.criterion1.length > MAX_FILTER_VALUE_LENGTH) return undefined;
      const hasCriterion2 = Object.prototype.hasOwnProperty.call(record, "criterion2");
      if (hasCriterion2 && (typeof record.criterion2 !== "string" || record.criterion2.length > MAX_FILTER_VALUE_LENGTH)) return undefined;
      const hasOperator = Object.prototype.hasOwnProperty.call(record, "operator");
      if (hasOperator && record.operator !== "And" && record.operator !== "Or") return undefined;
      return {
        kind: "table_filter_state",
        state: {
          columnIndex: rule.columnIndex,
          state: "filtered",
          criterion1: record.criterion1,
          ...(hasCriterion2 ? { criterion2: record.criterion2 as string } : {}),
          operator: hasOperator ? record.operator as "And" | "Or" : "And",
        },
      };
    }
    case "sheet_exists": return { kind: "sheet_exists", exists: value.isNullObject !== true };
    case "chart_exists": {
      const items = value.items;
      if (!Array.isArray(items)) return undefined;
      return {
        kind: "chart_exists",
        exists: items.some((chart) => {
          if (chart === null || typeof chart !== "object") return false;
          const title = (chart as Record<string, unknown>).title;
          return title !== null && typeof title === "object" && (title as Record<string, unknown>).text === rule.title;
        }),
      };
    }
    case "comment_exists": {
      const candidates = request.commentCandidates;
      if (candidates === undefined || request.expectedAddress === undefined) return undefined;
      const expectedAddress = normalizedCommentAddress(request.expectedAddress);
      const match = candidates.find(({ location }) => typeof location.address === "string" && normalizedCommentAddress(location.address) === expectedAddress);
      return { kind: "comment_exists", exists: match !== undefined, text: match?.comment.content };
    }
    case "calculation_completed": return { kind: "calculation_completed", completed: value.calculationState === "Done" || value.calculationState === "done" };
  }
}

function failedVerification(): WorkbookPlanVerificationReport {
  return { ok: false, rules: [], counts: { verified: 0, mismatched: 0, unsupported: 0, error: 1 }, firstFailure: null };
}

function isOfficeCancellationRequested(signal: AbortSignal | undefined): boolean {
  try { return signal !== undefined && signal.aborted === true; } catch { return true; }
}

function requireLoad(target: unknown, fields: string | readonly string[]): void {
  if (target === null || typeof target !== "object") throw new Error("Office load owner unavailable");
  const load = (target as { readonly load?: unknown }).load;
  if (typeof load !== "function") throw new Error("Office load method unavailable");
  load.call(target, fields);
}

function stringArg(args: Readonly<Record<string, unknown>>, key: string): string {
  const value = args[key];
  if (typeof value !== "string") throw new Error(`Missing string argument ${key}`);
  return value;
}
function numberArg(args: Readonly<Record<string, unknown>>, key: string): number {
  const value = args[key];
  if (typeof value !== "number" || !Number.isSafeInteger(value)) throw new Error(`Missing numeric argument ${key}`);
  return value;
}
function recordArg(args: Readonly<Record<string, unknown>>, key: string): Record<string, unknown> {
  const value = args[key];
  if (value === null || typeof value !== "object" || Array.isArray(value)) throw new Error(`Missing object argument ${key}`);
  return value as Record<string, unknown>;
}
