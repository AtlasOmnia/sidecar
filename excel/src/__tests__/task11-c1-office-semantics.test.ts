import { describe, expect, it } from "vitest";
import {
  compileWorkbookPlan,
  type VerifiedWorkbookCompileContext,
  type WorkbookPlanPreconditionManifest,
} from "../workbookPlanCompiler";
import {
  BrowserIdempotencyStore,
  MemoryIdempotencyStorage,
  createTrustedOperationBoundary,
  executeAcceptedWorkbookPlan,
  prepareWorkbookPlanExecution,
} from "../workbookPlanExecutor";
import {
  createNativeWorkbookPlanAdapter,
  readNativeWorkbookPlanRuntimeContext,
} from "../workbookPlanOfficeAdapter";
import type {
  WorkbookPlanEffectGroup,
  WorkbookPlanEffectOperation,
  WorkbookPlanFirstGroupContext,
  WorkbookPlanGroupExecutionResult,
} from "../workbookPlanExecutor";
import type { WorkbookOperation, WorkbookPlan, VerificationRule } from "../generated/workbookPlanSchema";
import type { WorkbookPlanPolicyInput } from "../workbookPlanPolicy";

type AnyRecord = Record<string, any>;
type Effect = "write" | "format" | "structure" | "filter" | "calculate";

const PRIMITIVES = [
  "write_range",
  "autofill_formula",
  "format_range",
  "insert_table",
  "sort_table",
  "filter_table",
  "create_chart",
  "set_data_validation",
  "conditional_format_range",
  "freeze_panes",
  "create_sheet",
  "rename_sheet",
  "copy_sheet",
  "add_comment",
  "run_calculation",
] as const;

const IDS = {
  sessionId: "11111111-1111-4111-8111-111111111111",
  documentId: "22222222-2222-4222-8222-222222222222",
  turnId: "33333333-3333-4333-8333-333333333333",
};

function dimensions(address: string): { rows: number; columns: number } {
  const local = address.split("!").at(-1)?.replace(/\$/g, "") ?? address;
  const [first, last = first] = local.split(":");
  const parse = (cell: string): { row: number; column: number } => {
    const match = /^([A-Z]+)([0-9]+)$/i.exec(cell);
    if (match === null) throw new Error(`strict fake only accepts A1 ranges: ${address}`);
    let column = 0;
    for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
    return { row: Number(match[2]), column };
  };
  const start = parse(first);
  const end = parse(last);
  return { rows: end.row - start.row + 1, columns: end.column - start.column + 1 };
}

function asFields(fields: string | readonly string[] | undefined): string[] {
  if (fields === undefined) return [];
  const values = typeof fields === "string" ? fields.split(",") : [...fields];
  return values.map((field: string) => field.trim());
}

function createStrictFake(options: {
  readonly failOnNumberFormat?: boolean;
  readonly freshWorksheetProxies?: boolean;
  readonly initialComments?: readonly { address: string; content: string }[];
  readonly reportedCommentCount?: unknown;
  readonly syntheticCommentCount?: number;
  readonly pauseCommentScanSync?: { readonly pending: Promise<void> };
  readonly onMutationQueued?: () => void;
  readonly initialValues?: Readonly<Record<string, unknown>>;
  readonly initialNumberFormats?: Readonly<Record<string, string>>;
  readonly unavailable?: readonly ("validation" | "conditional")[];
} = {}) {
  const calls: Array<{ kind: string; args: unknown[] }> = [];
  const pendingLoads: Array<() => void> = [];
  const pendingMutations: Array<() => void> = [];
  const ranges = new Map<string, AnyRecord>();
  const canonicalAddresses = new WeakMap<object, string>();
  const commentsBySheet = new Map<string, AnyRecord[]>();
  const sheets = new Map<string, AnyRecord>();
  let syncCount = 0;
  let unsynchronizedMutationCommands = 0;
  let committedMutationCommands = 0;
  let hiddenFinalSyncCount = 0;
  let commentCountReadQueued = false;
  let commentScanSyncStarted = false;
  let runCount = 0;
  let verificationCalls = 0;
  let nextCommentId = 1;

  const queueLoad = (commit: () => void): void => { pendingLoads.push(commit); };
  const queueMutation = (call: { kind: string; args: unknown[] }, commit?: () => void): void => {
    unsynchronizedMutationCommands += 1;
    calls.push(call);
    if (commit !== undefined) pendingMutations.push(commit);
    options.onMutationQueued?.();
  };

  const defineLoaded = (target: AnyRecord, key: string, initial: unknown, onSet?: (value: unknown) => void): void => {
    let value = initial;
    let loaded = false;
    Object.defineProperty(target, key, {
      configurable: true,
      enumerable: true,
      get: () => {
        if (!loaded) throw new Error(`read before load+sync: ${key}`);
        return value;
      },
      set: (next: unknown) => { if (onSet === undefined) value = next; else onSet(next); },
    });
    Object.defineProperty(target, `__load_${key}`, {
      value: () => queueLoad(() => { loaded = true; }),
      enumerable: false,
    });
  };

  const makeIdentityProbe = (exists: () => boolean): AnyRecord => {
    let loaded = false;
    let isNullObject = true;
    const probe: AnyRecord = {};
    Object.defineProperty(probe, "isNullObject", {
      enumerable: true,
      get: () => {
        if (!loaded) throw new Error("read before identity load+sync");
        return isNullObject;
      },
    });
    probe.load = (fields: string | readonly string[]) => {
      if (asFields(fields).join(",") !== "isNullObject") throw new Error("strict fake rejects identity load");
      loaded = false;
      queueLoad(() => {
        isNullObject = !exists();
        loaded = true;
      });
      return probe;
    };
    return probe;
  };

  const makeRange = (sheetName: string, address: string): AnyRecord => {
    const key = `${sheetName}!${address}`;
    const existing = ranges.get(key);
    if (existing !== undefined) return existing;
    const size = dimensions(address);
    const fullAddress = `${sheetName}!${address}`;
    const range: AnyRecord = {};
    canonicalAddresses.set(range, fullAddress);
    const initialValues = Array.from({ length: size.rows }, () => Array(size.columns).fill(null));
    if (Object.prototype.hasOwnProperty.call(options.initialValues ?? {}, fullAddress)) {
      initialValues[0]![0] = options.initialValues![fullAddress];
    }
    defineLoaded(range, "address", fullAddress);
    defineLoaded(range, "rowCount", size.rows);
    defineLoaded(range, "columnCount", size.columns);
    const initialFormulas = Array.from({ length: size.rows }, () => Array(size.columns).fill(""));
    defineLoaded(range, "values", initialValues, (value) => queueMutation(
      { kind: "values", args: [fullAddress, value] },
      () => { initialValues.splice(0, initialValues.length, ...(value as unknown[][])); },
    ));
    defineLoaded(range, "formulas", initialFormulas, (value) => queueMutation(
      { kind: "formulas", args: [fullAddress, value] },
      () => { initialFormulas.splice(0, initialFormulas.length, ...(value as unknown[][])); },
    ));
    Object.defineProperty(range, "__setFormulaState", {
      value: (value: readonly (readonly unknown[])[]) => {
        initialFormulas.splice(0, initialFormulas.length, ...value.map((row) => [...row]));
      },
      enumerable: false,
    });
    range.load = (fields: string | readonly string[]) => {
      for (const field of asFields(fields)) {
        const base = field.split("/")[0];
        const loader = range[`__load_${base}`];
        if (typeof loader !== "function") throw new Error(`strict fake rejects unknown load field ${field}`);
        loader();
      }
      return range;
    };
    let numberFormatLoaded = false;
    let numberFormatValue = Array.from({ length: size.rows }, () => Array(size.columns).fill(
      options.initialNumberFormats?.[fullAddress] ?? "General",
    ));
    Object.defineProperty(range, "numberFormat", {
      configurable: true,
      enumerable: true,
      get: () => {
        if (!numberFormatLoaded) throw new Error("read before load+sync: numberFormat");
        return numberFormatValue;
      },
      set: (value: unknown) => {
        if (options.failOnNumberFormat === true && address === "B1") throw new Error("fake Office InvalidArgument on numberFormat");
        if (!Array.isArray(value) || value.length !== size.rows || value.some((row) => !Array.isArray(row) || row.length !== size.columns)) {
          throw new Error(`strict fake requires dimensioned numberFormat for ${fullAddress}`);
        }
        queueMutation(
          { kind: "numberFormat", args: [fullAddress, value] },
          () => { numberFormatValue = value; },
        );
      },
    });
    Object.defineProperty(range, "__load_numberFormat", { value: () => queueLoad(() => { numberFormatLoaded = true; }) });

    const format: AnyRecord = {};
    const formatState: AnyRecord = { columnWidth: 100, rowHeight: 20 };
    let formatLoaded = false;
    const font: AnyRecord = {};
    const fontState: AnyRecord = {};
    let fontLoaded = false;
    font.load = (fields: string | readonly string[]) => {
      if (asFields(fields).some((field) => !["bold", "italic", "color", "name", "size", "underline"].includes(field))) throw new Error("strict fake rejects font load");
      queueLoad(() => { fontLoaded = true; });
      return font;
    };
    for (const property of ["bold", "italic", "color", "name", "size", "underline"]) {
      Object.defineProperty(font, property, {
        configurable: true,
        enumerable: true,
        set: (value: unknown) => queueMutation(
          { kind: `font.${property}`, args: [fullAddress, value] },
          () => { fontState[property] = value; },
        ),
        get: () => { if (!fontLoaded) throw new Error(`read before load+sync: font.${property}`); return fontState[property]; },
      });
    }
    const fill: AnyRecord = {};
    const fillState: AnyRecord = { color: "" };
    let fillLoaded = false;
    fill.load = (fields: string | readonly string[]) => {
      if (asFields(fields).join(",") !== "color") throw new Error("strict fake rejects fill load");
      queueLoad(() => { fillLoaded = true; });
      return fill;
    };
    Object.defineProperty(fill, "color", {
      enumerable: true,
      set: (value: unknown) => queueMutation(
        { kind: "fill.color", args: [fullAddress, value] },
        () => { fillState.color = value; },
      ),
      get: () => { if (!fillLoaded) throw new Error("read before load+sync: fill.color"); return fillState.color; },
    });
    const borderNames = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
    const borderState = new Map<string, AnyRecord>(borderNames.map((name) => [name, {}]));
    let bordersLoaded = false;
    const borders: AnyRecord = {
      load: (fields: string | readonly string[]) => {
        if (asFields(fields).some((field) => !["items/sideIndex", "items/style", "items/color", "items/weight"].includes(field))) throw new Error("strict fake rejects border load");
        queueLoad(() => { bordersLoaded = true; });
        return borders;
      },
      getItem: (name: string) => {
        if (!borderNames.includes(name)) throw new Error(`strict fake rejects invalid border ${name}`);
        const state = borderState.get(name)!;
        const border: AnyRecord = {};
        for (const property of ["style", "color", "weight"]) {
          Object.defineProperty(border, property, {
            enumerable: true,
            set: (value: unknown) => queueMutation(
              { kind: `border.${property}`, args: [fullAddress, name, value] },
              () => { state[property] = value; },
            ),
            get: () => { if (!bordersLoaded) throw new Error(`read before load+sync: border.${property}`); return state[property]; },
          });
        }
        return border;
      },
    };
    Object.defineProperty(borders, "items", {
      enumerable: true,
      get: () => {
        if (!bordersLoaded) throw new Error("read before load+sync: borders.items");
        return borderNames.map((sideIndex) => ({ sideIndex, ...borderState.get(sideIndex) }));
      },
    });
    for (const property of ["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation", "columnWidth", "rowHeight"]) {
      Object.defineProperty(format, property, {
        enumerable: true,
        set: (value: unknown) => queueMutation(
          { kind: `format.${property}`, args: [fullAddress, value] },
          () => { formatState[property] = value; },
        ),
        get: () => { if (!formatLoaded) throw new Error(`read before load+sync: format.${property}`); return formatState[property]; },
      });
    }
    format.load = (fields: string | readonly string[]) => {
      if (asFields(fields).some((field) => !["horizontalAlignment", "verticalAlignment", "wrapText", "indentLevel", "textOrientation", "columnWidth", "rowHeight"].includes(field))) throw new Error("strict fake rejects format load");
      queueLoad(() => { formatLoaded = true; });
      return format;
    };
    format.font = font;
    format.fill = fill;
    format.borders = borders;
    format.autofitColumns = () => queueMutation({ kind: "autofitColumns", args: [fullAddress] }, () => { formatState.columnWidth = 100; });
    format.autofitRows = () => queueMutation({ kind: "autofitRows", args: [fullAddress] }, () => { formatState.rowHeight = 20; });
    range.format = format;

    const validation: AnyRecord = {};
    let validationLoaded = false;
    let validationRule: unknown = undefined;
    let validationPrompt: unknown = undefined;
    validation.load = (fields: string | readonly string[]) => {
      if (asFields(fields).some((field) => !["type", "rule", "prompt"].includes(field))) throw new Error("strict fake rejects validation load");
      queueLoad(() => { validationLoaded = true; });
      return validation;
    };
    Object.defineProperty(validation, "type", {
      enumerable: true,
      get: () => { if (!validationLoaded) throw new Error("read before load+sync: validation.type"); return validationRule === undefined ? "None" : "List"; },
    });
    Object.defineProperty(validation, "rule", {
      set: (value: AnyRecord) => {
        if (Object.keys(value).length !== 1 || value.list === undefined || Object.keys(value.list).sort().join(",") !== "inCellDropDown,source") {
          throw new Error("strict fake rejects invented data validation rule shape");
        }
        queueMutation({ kind: "validation.rule", args: [value] }, () => { validationRule = value; });
      },
      get: () => { if (!validationLoaded) throw new Error("read before load+sync: validation.rule"); return validationRule; },
    });
    Object.defineProperty(validation, "prompt", {
      set: (value: AnyRecord) => {
        if (Object.keys(value).sort().join(",") !== "message,showPrompt,title") throw new Error("strict fake rejects prompt shape");
        queueMutation({ kind: "validation.prompt", args: [value] }, () => { validationPrompt = value; });
      },
      get: () => { if (!validationLoaded) throw new Error("read before load+sync: validation.prompt"); return validationPrompt; },
    });
    range.dataValidation = options.unavailable?.includes("validation") === true ? undefined : validation;

    let conditionalSequence = 0;
    let conditionalItemsLoaded = false;
    const conditionalItems: AnyRecord[] = [];
    const makeConditional = (type: "CellValue" | "ContainsText"): AnyRecord => {
      const state: AnyRecord = { rule: undefined, fillColor: undefined, fontColor: undefined };
      const formatObject: AnyRecord = { fill: {}, font: {} };
      Object.defineProperty(formatObject.fill, "color", {
        enumerable: true,
        set: (value: unknown) => queueMutation(
          { kind: `${type}.format.fill`, args: [value] },
          () => { state.fillColor = value; },
        ),
        get: () => {
          if (!conditionalItemsLoaded) throw new Error(`read before load+sync: ${type}.format.fill`);
          return state.fillColor;
        },
      });
      Object.defineProperty(formatObject.font, "color", {
        enumerable: true,
        set: (value: unknown) => queueMutation(
          { kind: `${type}.format.font`, args: [value] },
          () => { state.fontColor = value; },
        ),
        get: () => {
          if (!conditionalItemsLoaded) throw new Error(`read before load+sync: ${type}.format.font`);
          return state.fontColor;
        },
      });
      const ruleTarget: AnyRecord = {};
      Object.defineProperty(ruleTarget, "rule", {
        enumerable: true,
        set: (value: AnyRecord) => {
          const expectedKeys = type === "CellValue" ? "formula1,operator" : "operator,text";
          if (Object.keys(value).sort().join(",") !== expectedKeys) throw new Error(`strict fake rejects ${type} rule shape`);
          const expectedOperator = type === "CellValue" ? "GreaterThan" : "Contains";
          if (value.operator !== expectedOperator) throw new Error(`strict fake rejects ${type} operator`);
          queueMutation({ kind: `${type}.rule`, args: [value] }, () => { state.rule = value; });
        },
        get: () => {
          if (!conditionalItemsLoaded) throw new Error(`read before load+sync: ${type}.rule`);
          return state.rule;
        },
      });
      ruleTarget.format = formatObject;
      const result: AnyRecord = {
        id: `strict-cf-${conditionalSequence += 1}`,
        type,
        ...(type === "CellValue" ? { cellValue: ruleTarget } : { textComparison: ruleTarget }),
      };
      Object.defineProperty(result, "format", { get: () => { throw new Error("top-level conditional format mutation is forbidden"); } });
      return result;
    };
    const conditionalFormats: AnyRecord = {
      load: (fields: string | readonly string[]) => {
        const requested = asFields(fields);
        const identityOnly = requested.join(",") === "items/id";
        const supportedDetails = requested.every((field) => [
          "items/id", "items/type",
          "items/cellValue/rule", "items/cellValue/format/fill/color", "items/cellValue/format/font/color",
          "items/textComparison/rule", "items/textComparison/format/fill/color", "items/textComparison/format/font/color",
        ].includes(field));
        if (!identityOnly && !supportedDetails) throw new Error("strict fake rejects conditional readback load");
        conditionalItemsLoaded = false;
        queueLoad(() => { conditionalItemsLoaded = true; });
      },
      add: (type: string) => {
        if (type !== "CellValue" && type !== "ContainsText") throw new Error(`strict fake rejects conditional type ${type}`);
        const created = makeConditional(type);
        queueMutation({ kind: "conditional.add", args: [type] }, () => { conditionalItems.push(created); });
        return created;
      },
    };
    Object.defineProperty(conditionalFormats, "items", {
      enumerable: true,
      get: () => {
        if (!conditionalItemsLoaded) throw new Error("read before load+sync: conditionalFormats.items");
        return conditionalItems;
      },
    });
    range.conditionalFormats = options.unavailable?.includes("conditional") === true ? undefined : conditionalFormats;
    range.autoFill = (destination: AnyRecord, autoFillType?: string) => {
      const destinationAddress = destination === undefined ? undefined : canonicalAddresses.get(destination);
      if (destinationAddress === undefined || destination === range) throw new Error("strict fake requires source-driven destination range");
      if (autoFillType !== undefined) throw new Error("strict fake rejects unexpected autofill type argument");
      queueMutation(
        { kind: "autoFill", args: [fullAddress, destinationAddress] },
        () => {
          const sourceFormula = initialFormulas[0]?.[0];
          const setFormulaState = destination.__setFormulaState;
          if (sourceFormula !== "=A1+B1" || typeof setFormulaState !== "function") return;
          if (destinationAddress === "Sheet1!D1:F1") setFormulaState([["=A1+B1", "=B1+C1", "=C1+D1"]]);
          if (destinationAddress === "Sheet1!D1:D3") setFormulaState([["=A1+B1"], ["=A2+B2"], ["=A3+B3"]]);
        },
      );
    };
    Object.defineProperty(range, "fillDown", { get: () => { throw new Error("invented target fillDown was accessed"); } });
    ranges.set(key, range);
    return range;
  };

  const makeComment = (sheetName: string, address: string, content: string): AnyRecord => {
    const comment: AnyRecord = {};
    const location = makeRange(sheetName, address);
    comment.load = (fields: string | readonly string[]) => {
      for (const field of asFields(fields)) {
        if (field !== "content") throw new Error(`strict fake rejects comment load ${field}`);
        calls.push({ kind: "comment.content.load", args: [sheetName, address] });
        comment.__load_content();
      }
      return comment;
    };
    comment.getLocation = () => {
      calls.push({ kind: "comment.getLocation", args: [sheetName, address] });
      return location;
    };
    Object.defineProperty(comment, "content", {
      configurable: true,
      enumerable: true,
      get: () => {
        if (comment.__contentLoaded !== true) throw new Error("read before load+sync: comment.content");
        return comment.__contentValue;
      },
      set: (value: unknown) => {
        if (typeof value !== "string") throw new Error("comment content must be string");
        queueMutation(
          { kind: "comment.update", args: [sheetName, address, value] },
          () => {
            comment.__contentValue = value;
            comment.__contentLoaded = true;
          },
        );
      },
    });
    comment.__contentValue = content;
    comment.__contentLoaded = false;
    Object.defineProperty(comment, "__load_content", { value: () => queueLoad(() => { comment.__contentLoaded = true; }) });
    Object.defineProperty(comment, "id", { value: `comment-${nextCommentId++}` });
    return comment;
  };

  const makeComments = (sheetName: string): AnyRecord => {
    let items: AnyRecord[] = commentsBySheet.get(sheetName) ?? [];
    let loaded = false;
    let countLoaded = false;
    let countValue: unknown;
    const countResult: AnyRecord = {};
    Object.defineProperty(countResult, "value", {
      enumerable: true,
      get: () => {
        if (!countLoaded) throw new Error("read before load+sync: comments.getCount().value");
        return countValue;
      },
    });
    const comments: AnyRecord = {};
    Object.defineProperty(comments, "items", {
      enumerable: true,
      get: () => {
        if (!loaded) throw new Error("read before load+sync: comments.items");
        return items;
      },
    });
    comments.load = (fields: string | readonly string[]) => {
      const selected = asFields(fields);
      if (selected.length !== 1 || selected[0] !== "items") throw new Error("strict fake requires bounded comments.load(items)");
      if (!countLoaded) throw new Error("comments.load(items) requires loaded count");
      calls.push({ kind: "comments.load.items", args: [sheetName] });
      queueLoad(() => { loaded = true; items = commentsBySheet.get(sheetName) ?? []; });
      return comments;
    };
    comments.getCount = () => {
      commentCountReadQueued = true;
      calls.push({ kind: "comments.getCount", args: [sheetName] });
      queueLoad(() => {
        countValue = options.reportedCommentCount === undefined
          ? (commentsBySheet.get(sheetName) ?? []).length
          : options.reportedCommentCount;
        countLoaded = true;
      });
      return countResult;
    };
    comments.add = (range: AnyRecord, content: string) => {
      const canonicalAddress = canonicalAddresses.get(range);
      if (canonicalAddress === undefined) throw new Error("comments.add received an unknown range");
      const address = canonicalAddress.split("!").at(-1) ?? "";
      const comment = makeComment(sheetName, address, content);
      queueMutation(
        { kind: "comment.add", args: [canonicalAddress, content] },
        () => {
          items.push(comment);
          commentsBySheet.set(sheetName, items);
        },
      );
      return comment;
    };
    Object.defineProperty(comments, "getItemByCell", { get: () => { throw new Error("getItemByCell is throwing and forbidden"); } });
    return comments;
  };

  const filterStates = new Map<string, AnyRecord>();
  const makeFilter = (tableName: string, columnIndex: number): AnyRecord => {
    const filter: AnyRecord = {};
    const filterKey = `${tableName}:${columnIndex}`;
    let criteria: AnyRecord = filterStates.get(filterKey) ?? { filterOn: "None" };
    let criteriaLoaded = false;
    Object.defineProperty(filter, "criteria", {
      get: () => {
        if (!criteriaLoaded) throw new Error("read before load+sync: filter.criteria");
        return criteria;
      },
    });
    filter.load = (fields: string | readonly string[]) => {
      if (asFields(fields).join(",") !== "criteria") throw new Error("strict fake rejects unknown filter load");
      queueLoad(() => { criteriaLoaded = true; });
      return filter;
    };
    filter.applyValuesFilter = (values: readonly string[]) => {
      if (!Array.isArray(values) || values.some((value) => typeof value !== "string")) throw new Error("invalid values filter");
      const nextCriteria = { filterOn: "Values", values: [...values] };
      queueMutation(
        { kind: "filter.values", args: [tableName, columnIndex, values] },
        () => { criteria = nextCriteria; filterStates.set(filterKey, criteria); },
      );
    };
    filter.applyCustomFilter = (...args: unknown[]) => {
      if (args.length < 1 || args.length > 3 || typeof args[0] !== "string" || args.slice(1).some((value) => value !== undefined && typeof value !== "string")) throw new Error("invalid custom filter arguments");
      const nextCriteria = { filterOn: "Custom", criterion1: args[0], ...(args[1] === undefined ? {} : { criterion2: args[1] }), ...(args[2] === undefined ? {} : { operator: args[2] }) };
      queueMutation(
        { kind: "filter.custom", args: [tableName, columnIndex, ...args] },
        () => { criteria = nextCriteria; filterStates.set(filterKey, criteria); },
      );
    };
    filter.clear = () => {
      queueMutation(
        { kind: "filter.clear", args: [tableName, columnIndex] },
        () => { criteria = { filterOn: "None" }; filterStates.set(filterKey, criteria); },
      );
    };
    Object.defineProperty(filter, "clearCriteria", { get: () => { throw new Error("invented clearCriteria was accessed"); } });
    return filter;
  };

  const gateTable = (
    target: AnyRecord,
    initial: { readonly isNullObject: boolean; readonly name: string; readonly showHeaders: boolean; readonly showTotals: boolean },
    onNameCommitted?: (name: string) => void,
  ): void => {
    const loadedFields = new Set<string>();
    const state = { ...initial };
    for (const [key, value] of Object.entries(state)) {
      Object.defineProperty(target, key, {
        configurable: true,
        enumerable: true,
        get: () => {
          if (!loadedFields.has(key)) throw new Error(`read before table load+sync: ${key}`);
          return state[key as keyof typeof state];
        },
        set: (next: unknown) => {
          if (key !== "name" || typeof next !== "string") throw new Error(`strict fake rejects table ${key} mutation`);
          queueMutation(
            { kind: "table.name", args: [next] },
            () => {
              state.name = next;
              onNameCommitted?.(next);
            },
          );
        },
      });
      Object.defineProperty(target, `__load_table_${key}`, { value: () => queueLoad(() => { loadedFields.add(key); }) });
    }
    target.load = (fields: string | readonly string[]) => {
      const requested = asFields(fields);
      if (requested.some((field) => !["isNullObject", "name", "showHeaders", "showTotals"].includes(field))) throw new Error("strict fake rejects table load");
      for (const field of requested) {
        loadedFields.delete(field);
        target[`__load_table_${field}`]();
      }
      return target;
    };
  };

  const extraTables = new Map<string, AnyRecord>();
  let tableSortLoaded = false;
  let tableSortFields: readonly AnyRecord[] = [];
  let tableBodyLoaded = false;
  let tableBodyReadCount = 0;
  let tableBodyValues: unknown[][] = [[10, 30, 50], [20, 10, 40], [30, 20, 60]];
  const tableBodyRange: AnyRecord = {};
  Object.defineProperty(tableBodyRange, "values", {
    get: () => {
      if (!tableBodyLoaded) throw new Error("read before load+sync: table data body values");
      tableBodyReadCount += 1;
      return tableBodyValues;
    },
  });
  tableBodyRange.load = (fields: string | readonly string[]) => {
    if (asFields(fields).join(",") !== "values") throw new Error("strict fake rejects table data-body load");
    tableBodyLoaded = false;
    queueLoad(() => { tableBodyLoaded = true; });
  };
  const table = {
    isNullObject: false,
    name: "Table1",
    showHeaders: true,
    showTotals: false,
    load: () => undefined,
    range: makeRange("Sheet1", "A1:C4"),
    sort: {
      load: (fields: string | readonly string[]) => {
        if (asFields(fields).join(",") !== "fields") throw new Error("strict fake rejects table sort load");
        tableSortLoaded = false;
        queueLoad(() => { tableSortLoaded = true; });
      },
      get fields() {
        if (!tableSortLoaded) throw new Error("read before load+sync: table.sort.fields");
        return tableSortFields;
      },
      apply: (fields: readonly AnyRecord[]) => {
        if (!Array.isArray(fields) || fields.length !== 1 || Object.keys(fields[0]).sort().join(",") !== "ascending,key" || typeof fields[0].key !== "number" || typeof fields[0].ascending !== "boolean") {
          throw new Error("strict fake rejects non-SortField table sort");
        }
        queueMutation(
          { kind: "table.sort", args: [fields] },
          () => {
            tableSortFields = fields.map((field) => ({ ...field }));
            const key = fields[0]!.key;
            const ascending = fields[0]!.ascending;
            tableBodyValues = [...tableBodyValues].sort((left, right) => {
              const leftKey = left[key];
              const rightKey = right[key];
              if (typeof leftKey !== "number" || typeof rightKey !== "number") return 0;
              const comparison = leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
              return ascending ? comparison : -comparison;
            });
          },
        );
      },
    },
    columns: {
      getItemAt: (index: number) => {
        if (!Number.isInteger(index)) throw new Error("invalid table column index");
        return { filter: makeFilter("Table1", index) };
      },
    },
    getRange: () => table.range,
    getDataBodyRange: () => tableBodyRange,
  } as AnyRecord;
  gateTable(table, { isNullObject: false, name: "Table1", showHeaders: true, showTotals: false });

  const makeSheet = (name: string, fresh = false, register = true): AnyRecord => {
    const existing = sheets.get(name);
    if (!fresh && existing !== undefined) return existing;
    let assignedName = name;
    let sheetNameLoaded = false;
    let frozenRows = 0;
    let frozenColumns = 0;
    const chartsByName = new Map<string, AnyRecord>();
    let freezeLocationLoaded = false;
    const freezeLocation: AnyRecord = {};
    for (const [key, value] of [["isNullObject", false], ["address", "A1"], ["rowCount", 2], ["columnCount", 3]] as const) {
      Object.defineProperty(freezeLocation, key, {
        enumerable: true,
        get: () => {
          if (!freezeLocationLoaded) throw new Error(`read before load+sync: freeze.${key}`);
          if (key === "isNullObject") return frozenRows === 0 && frozenColumns === 0;
          if (key === "rowCount") return frozenRows;
          if (key === "columnCount") return frozenColumns;
          return value;
        },
      });
    }
    freezeLocation.load = (fields: string | readonly string[]) => {
      if (asFields(fields).join(",") !== "isNullObject,address,rowCount,columnCount") throw new Error("strict fake rejects freeze load");
      queueLoad(() => { freezeLocationLoaded = true; });
      return freezeLocation;
    };
    const makeChart = (type: string, sourceAddress: string, onNameCommitted: (name: string, chart: AnyRecord) => void): AnyRecord => {
      let chartLoaded = false;
      let titleLoaded = false;
      const state = { name: "", chartType: type, title: "", left: 10, top: 10, width: 200, height: 100 };
      const title: AnyRecord = {};
      Object.defineProperty(title, "text", {
        enumerable: true,
        get: () => {
          if (!titleLoaded) throw new Error("read before ChartTitle load+sync");
          return state.title;
        },
        set: (value: string) => queueMutation(
          { kind: "chart.title", args: [value] },
          () => { state.title = value; },
        ),
      });
      const chart: AnyRecord = { title };
      for (const key of ["isNullObject", "name", "chartType", "left", "top", "width", "height"] as const) {
        Object.defineProperty(chart, key, {
          configurable: true,
          enumerable: true,
          get: () => {
            if (!chartLoaded) throw new Error(`read before Chart load+sync: ${key}`);
            return key === "isNullObject" ? false : state[key];
          },
          set: (value: unknown) => {
            if (key === "name" && typeof value === "string") {
              queueMutation({ kind: "chart.name", args: [value] }, () => {
                state.name = value;
                onNameCommitted(value, chart);
              });
              return;
            }
            if (key === "chartType" && typeof value === "string") {
              queueMutation({ kind: "chart.type", args: [value] }, () => { state.chartType = value; });
              return;
            }
            throw new Error(`strict fake rejects chart ${key} mutation`);
          },
        });
      }
      chart.load = (fields: string | readonly string[]) => {
        if (asFields(fields).join(",") !== "isNullObject,name,chartType,title/text,left,top,width,height") throw new Error("strict fake rejects chart load");
        chartLoaded = false;
        titleLoaded = false;
        queueLoad(() => { chartLoaded = true; titleLoaded = true; });
        return chart;
      };
      Object.defineProperty(chart, "__load_chart_title", { value: () => { titleLoaded = true; } });
      Object.defineProperty(chart, "__chartSourceAddress", { value: sourceAddress });
      return chart;
    };
    let chartsLoaded = false;
    const charts: AnyRecord = {
      add: (type: string, source: AnyRecord, seriesBy?: string) => {
        const allowed = new Set(["ColumnClustered", "Line", "Pie"]);
        const sourceAddress = canonicalAddresses.get(source);
        if (!allowed.has(type) || sourceAddress === undefined || seriesBy !== "Auto") throw new Error("invalid chart add shape");
        const chart = makeChart(type, sourceAddress, (chartName, chartObject) => {
          if (chartName.length > 0) chartsByName.set(chartName, chartObject);
        });
        queueMutation({ kind: "chart.add", args: [type, sourceAddress, seriesBy, chart] });
        return chart;
      },
      load: (fields: string | readonly string[]) => {
        if (asFields(fields).join(",") !== "items/title/text") throw new Error("invalid chart load");
        chartsLoaded = false;
        queueLoad(() => {
          chartsLoaded = true;
          for (const chart of chartsByName.values()) chart.__load_chart_title();
        });
        return charts;
      },
      get items() {
        if (!chartsLoaded) throw new Error("read before Charts load+sync");
        return [...chartsByName.values()];
      },
      getItemOrNullObject: (chartName: string) => chartsByName.get(chartName) ?? makeIdentityProbe(() => chartsByName.has(chartName)),
    };
    const sheet: AnyRecord = {
      get name() {
        if (!sheetNameLoaded) throw new Error("read before worksheet name load+sync");
        return assignedName;
      },
      load: (fields: string | readonly string[]) => {
        if (asFields(fields).join(",") !== "name") throw new Error("strict fake rejects worksheet load");
        sheetNameLoaded = false;
        queueLoad(() => { sheetNameLoaded = true; });
        return sheet;
      },
      set name(value: string) {
        queueMutation(
          { kind: "sheet.rename", args: [name, value] },
          () => {
            assignedName = value;
            const currentKey = Array.from(sheets.entries()).find(([, candidate]) => candidate === sheet)?.[0];
            if (currentKey !== undefined) {
              sheets.delete(currentKey);
              sheets.set(value, sheet);
            }
          },
        );
      },
      getRange: (address: string) => makeRange(name, address),
      getRanges: (address: string) => {
        const areas: AnyRecord = {};
        defineLoaded(areas, "address", `${name}!${address}`);
        defineLoaded(areas, "areaCount", 1);
        areas.load = (fields: string | readonly string[]) => {
          for (const field of asFields(fields)) {
            const loader = areas[`__load_${field}`];
            if (typeof loader !== "function") throw new Error(`strict fake rejects unknown range-area load field ${field}`);
            loader();
          }
        };
        return areas;
      },
      tables: { add: (range: AnyRecord, hasHeaders: boolean) => {
        const canonicalAddress = canonicalAddresses.get(range);
        if (canonicalAddress === undefined || typeof hasHeaders !== "boolean") throw new Error("invalid table add arguments");
        const created = {
          isNullObject: false,
          name: "",
          showHeaders: hasHeaders,
          showTotals: false,
          range,
          sort: table.sort,
          columns: table.columns,
          getRange: () => range,
        } as AnyRecord;
        gateTable(created, { isNullObject: false, name: "", showHeaders: hasHeaders, showTotals: false }, (name) => {
          if (name.length > 0) extraTables.set(name, created);
        });
        queueMutation(
          { kind: "table.add", args: [canonicalAddress, hasHeaders] },
        );
        return created;
      } },
 charts,
 comments: makeComments(name),
      getUsedRangeOrNullObject: () => ({
        isNullObject: true,
        load: (fields: string | readonly string[]) => {
          if (asFields(fields).sort().join(",") !== "address,isNullObject") throw new Error("strict fake rejects used-range load");
        },
      }),
      freezePanes: {
        freezeRows: (count: number) => { if (!Number.isSafeInteger(count) || count <= 0) throw new Error("invalid freeze rows"); queueMutation({ kind: "freezeRows", args: [count] }, () => { frozenRows = count; }); },
        freezeColumns: (count: number) => { if (!Number.isSafeInteger(count) || count <= 0) throw new Error("invalid freeze columns"); queueMutation({ kind: "freezeColumns", args: [count] }, () => { frozenColumns = count; }); },
        unfreeze: () => queueMutation({ kind: "unfreeze", args: [] }, () => { frozenRows = 0; frozenColumns = 0; }),
        getLocationOrNullObject: () => freezeLocation,
      },
      copy: (position: string) => {
        if (!["None", "Before", "After", "Beginning", "End"].includes(position)) throw new Error("invalid worksheet copy position");
        const copied = makeSheet(`${name} Copy`, false, false);
        queueMutation(
          { kind: "sheet.copy", args: [position] },
          () => {
            sheets.set(`${name} Copy`, copied);
            commentsBySheet.set(`${name} Copy`, commentsBySheet.get(`${name} Copy`) ?? []);
          },
        );
        return copied;
      },
    };
    if (register) {
      if (existing === undefined) sheets.set(name, sheet);
      commentsBySheet.set(name, commentsBySheet.get(name) ?? []);
    }
    return sheet;
  };

  const initialComments = options.initialComments ?? [];
  const sheet1 = makeSheet("Sheet1");
  for (const item of initialComments) {
    const list = commentsBySheet.get("Sheet1")!;
    list.push(makeComment("Sheet1", item.address, item.content));
  }
  for (let index = 0; index < (options.syntheticCommentCount ?? 0); index += 1) {
    const list = commentsBySheet.get("Sheet1")!;
    list.push(makeComment("Sheet1", `Z${index + 1}`, `c${index}`));
  }

  let worksheetsLoaded = false;
  let tablesLoaded = false;
  let calculationLoaded = false;
  let calculationState = "Done";
  const workbook: AnyRecord = {
    worksheets: {
      getItem: (name: string) => makeSheet(name, options.freshWorksheetProxies === true),
      getItemOrNullObject: (name: string) => makeIdentityProbe(() => sheets.has(name)),
      add: (name: string) => {
        const added = makeSheet(name, false, false);
        queueMutation(
          { kind: "sheet.add", args: [name] },
          () => {
            sheets.set(name, added);
            commentsBySheet.set(name, commentsBySheet.get(name) ?? []);
          },
        );
        return added;
      },
    },
    tables: {
      getItem: (name: string) => { if (name === "Table1") return table; const extra = extraTables.get(name); if (extra !== undefined) return extra; throw new Error(`unknown table ${name}`); },
      getItemOrNullObject: (name: string) => {
        if (name === "Table1") return table;
        const extra = extraTables.get(name);
        if (extra !== undefined) return extra;
        const probe = makeIdentityProbe(() => extraTables.has(name));
        probe.getRange = () => table.range;
        return probe;
      },
    },
    application: {
      calculate: (type: string) => {
        if (type !== "Full") throw new Error("invalid calculation type");
        queueMutation({ kind: "calculate", args: [type] }, () => { calculationState = "Done"; });
      },
      get calculationState() {
        if (!calculationLoaded) throw new Error("read before Application load+sync: calculationState");
        return calculationState;
      },
      load: (fields: string | readonly string[]) => {
        if (asFields(fields).join(",") !== "calculationState") throw new Error("invalid calculation load");
        calculationLoaded = false;
        queueLoad(() => { calculationLoaded = true; });
      },
    },
  };

  workbook.worksheets.load = (fields: string | readonly string[]) => {
    if (asFields(fields).join(",") !== "items/name") throw new Error("strict fake rejects worksheet load");
    worksheetsLoaded = false;
    for (const item of sheets.values()) item.load("name");
    queueLoad(() => { worksheetsLoaded = true; });
  };
  Object.defineProperty(workbook.worksheets, "items", { get: () => {
    if (!worksheetsLoaded) throw new Error("read before Worksheets load+sync");
    return [...sheets.values()];
  } });
  workbook.worksheets.getActiveWorksheet = () => sheet1;
  workbook.tables.load = (fields: string | readonly string[]) => {
    if (asFields(fields).join(",") !== "items/name") throw new Error("strict fake rejects table load");
    tablesLoaded = false;
    for (const item of [table, ...extraTables.values()]) item.load("name");
    queueLoad(() => { tablesLoaded = true; });
  };
  Object.defineProperty(workbook.tables, "items", { get: () => {
    if (!tablesLoaded) throw new Error("read before Tables load+sync");
    return [table, ...extraTables.values()];
  } });

  const context = {
    workbook,
    sync: async () => {
      const pauseCommentScan = commentCountReadQueued && options.pauseCommentScanSync !== undefined && !commentScanSyncStarted;
      commentCountReadQueued = false;
      syncCount += 1;
      if (pauseCommentScan) {
        commentScanSyncStarted = true;
        await options.pauseCommentScanSync!.pending;
      }
      const queued = pendingLoads.splice(0);
      for (const commit of queued) commit();
      const mutations = pendingMutations.splice(0);
      for (const commit of mutations) commit();
      committedMutationCommands += unsynchronizedMutationCommands;
      unsynchronizedMutationCommands = 0;
      calls.push({ kind: "sync", args: [syncCount] });
    },
  };

  const fake = {
    context,
    calls,
    sheet: sheet1,
    get syncCount() { return syncCount; },
    get unsynchronizedMutationCommands() { return unsynchronizedMutationCommands; },
    get committedMutationCommands() { return committedMutationCommands; },
    get hiddenFinalSyncCount() { return hiddenFinalSyncCount; },
    get commentScanSyncStarted() { return commentScanSyncStarted; },
    get runCount() { return runCount; },
    get verificationCalls() { return verificationCalls; },
    get sortBodyReadCount() { return tableBodyReadCount; },
    discardUnsynchronizedCommands: () => {
      pendingLoads.splice(0);
      pendingMutations.splice(0);
      unsynchronizedMutationCommands = 0;
    },
    markRun: () => { runCount += 1; },
    markVerification: () => { verificationCalls += 1; },
    setCommentContent: (address: string, content: string) => {
      const comment = (commentsBySheet.get("Sheet1") ?? []).find((candidate) => {
        const location = candidate.getLocation();
        return location.address === `Sheet1!${address}`;
      });
      if (comment === undefined) throw new Error(`unknown comment ${address}`);
      comment.__contentValue = content;
      comment.__contentLoaded = true;
    },
    loadAndSync: async (range: AnyRecord, fields: string | readonly string[]) => { range.load(fields); await context.sync(); },
    syncHiddenFinalCommands: async () => {
      if (unsynchronizedMutationCommands === 0) return;
      hiddenFinalSyncCount += 1;
      await context.sync();
    },
  };
  return fake;
}

async function withExcel<T>(fake: ReturnType<typeof createStrictFake>, task: () => Promise<T>): Promise<T> {
  const globals = globalThis as Record<string, unknown>;
  const previous = globals.Excel;
  globals.Excel = {
    run: async (callback: (context: unknown) => Promise<unknown>) => {
      fake.markRun();
      try {
        const result = await callback(fake.context);
        await fake.syncHiddenFinalCommands();
        return result;
      } catch (error) {
        fake.discardUnsynchronizedCommands();
        throw error;
      }
    },
  };
  try {
    return await task();
  } finally {
    if (previous === undefined) delete globals.Excel;
    else globals.Excel = previous;
  }
}

function compileContext(preconditionManifest?: WorkbookPlanPreconditionManifest): VerifiedWorkbookCompileContext {
  return {
    sheetNames: new Set(["Sheet1"]),
    tablesByName: new Map([["Table1", { canonicalName: "Table1", columnCount: 3, rowCount: 4, sheetName: "Sheet1" }]]),
    supportedPrimitiveNames: new Set(PRIMITIVES),
    supportedCapabilities: new Set(["excel_workbook_access", "workbook_plan_v1_operation_audit"]),
    chartSourceResolutionsBySheet: new Map([["Sheet1", new Map([["A1:C4", { areaCount: 1 }]])]]),
    ...(preconditionManifest === undefined ? {} : { preconditionManifest }),
  };
}

function semanticPlan(): WorkbookPlan {
  const operations: WorkbookOperation[] = [
    { kind: "write_values", sheet: "Sheet1", address: "A1:C2", values: [[1, 2, 3], [4, 5, 6]] },
    { kind: "write_formula", sheet: "Sheet1", address: "D1", formula: "=A1+B1" },
    { kind: "fill_formula", sheet: "Sheet1", sourceAddress: "D1", targetAddress: "D1:F1" },
    { kind: "fill_formula", sheet: "Sheet1", sourceAddress: "D1", targetAddress: "D1:D3" },
    { kind: "add_comment", sheet: "Sheet1", address: "E1", text: "first" },
    { kind: "add_comment", sheet: "Sheet1", address: "E1", text: "updated" },
    {
      kind: "format_range",
      sheet: "Sheet1",
      address: "A1:C2",
      format: {
        bold: true,
        italic: true,
        fontColor: "#111111",
        fillColor: "#eeeeee",
        numberFormat: "0.00",
        borders: true,
        borderColor: "#222222",
        borderWeight: "Medium",
        columnAutofit: true,
        rowAutofit: true,
        horizontalAlignment: "Center",
        verticalAlignment: "Center",
        fontName: "Aptos",
        fontSize: 12,
        underline: "Single",
        wrapText: true,
        indentLevel: 2,
        textOrientation: 15,
      },
    },
    { kind: "set_validation_list", sheet: "Sheet1", address: "F1:F2", source: "one,two", promptTitle: "Pick", promptMessage: "Choose one" },
    { kind: "conditional_format", sheet: "Sheet1", address: "A1:C2", ruleType: "cellValue", operator: "greaterThan", formula1: "0", fillColor: "#00ff00", fontColor: "#000000" },
    { kind: "conditional_format", sheet: "Sheet1", address: "A1:C2", ruleType: "containsText", operator: "contains", text: "ok", fillColor: "#00ff00", fontColor: "#000000" },
    { kind: "freeze_panes", sheet: "Sheet1", rows: 2, columns: 3 },
    { kind: "freeze_panes", sheet: "Sheet1", rows: 0, columns: 0 },
    { kind: "create_table", sheet: "Sheet1", address: "A1:C4", hasHeaders: true },
    { kind: "create_chart", sheet: "Sheet1", dataRange: "A1:C4", chartType: "ColumnClustered", title: "Revenue" },
    { kind: "create_sheet", name: "NewSheet" },
    { kind: "rename_sheet", oldName: "NewSheet", newName: "RenamedSheet" },
    { kind: "copy_sheet", sourceName: "RenamedSheet", newName: "CopiedSheet" },
    { kind: "sort_table", tableName: "Table1", columnIndex: 1, ascending: false },
    { kind: "filter_table", tableName: "Table1", columnIndex: 0, values: ["one", "two"] },
    { kind: "filter_table", tableName: "Table1", columnIndex: 1, criterion1: ">0", criterion2: "<10", operator: "And" },
    { kind: "reset_table_filter", tableName: "Table1", columnIndex: 0 },
    { kind: "calculate" },
  ];
  return { version: 1, outcome: "execute", summary: "C1 strict Office semantics", operations, verification: [{ kind: "calculation_completed" }] };
}

function compilerGroups(plan: WorkbookPlan): { primitives: readonly AnyRecord[]; groups: readonly WorkbookPlanEffectGroup[] } {
  const result = compileWorkbookPlan(plan, compileContext());
  expect(result.ok).toBe(true);
  if (!result.ok) throw new Error(JSON.stringify(result.error));
  const primitives = result.transaction.primitives as readonly AnyRecord[];
  const groups: WorkbookPlanEffectGroup[] = [];
  for (const primitive of primitives) {
    const item: WorkbookPlanEffectOperation = {
      operationId: `${String(primitive.operationIndex + 1).padStart(8, "0")}-1111-4111-8111-111111111111`,
      operationIndex: primitive.operationIndex,
      kind: primitive.kind,
      primitiveName: primitive.primitiveName,
      args: primitive.args,
    };
    const previous = groups.at(-1);
    if (previous !== undefined && previous.effect === primitive.effect) {
      groups[groups.length - 1] = { ...previous, operations: [...previous.operations, item] };
    } else {
      groups.push({ groupIndex: groups.length, effect: primitive.effect as Effect, operations: [item] });
    }
  }
  return { primitives, groups };
}

function compilerGroupForOperations(operations: readonly WorkbookOperation[]): WorkbookPlanEffectGroup {
  const compiled = compilerGroups({
    version: 1,
    outcome: "execute",
    summary: "C1 optional Office semantics",
    operations,
    verification: [],
  });
  expect(compiled.groups).toHaveLength(1);
  return compiled.groups[0]!;
}

async function executeCompilerOperations(
  fake: ReturnType<typeof createStrictFake>,
  operations: readonly WorkbookOperation[],
): Promise<WorkbookPlanGroupExecutionResult> {
  const adapter = createNativeWorkbookPlanAdapter();
  return withExcel(fake, () => adapter.executeGroup(compilerGroupForOperations(operations), 0));
}

async function loadMutationTargets(fake: ReturnType<typeof createStrictFake>): Promise<{ range: AnyRecord; comment: AnyRecord }> {
  const range = fake.sheet.getRange("A1");
  range.load(["values", "formulas", "numberFormat"]);
  const comments = fake.sheet.comments;
  comments.getCount();
  await fake.context.sync();
  comments.load("items");
  await fake.context.sync();
  const comment = comments.items[0];
  comment.load("content");
  await fake.context.sync();
  return { range, comment };
}

function preconditionManifest(): WorkbookPlanPreconditionManifest {
  return {
    version: 1,
    sheets: [{ name: "Sheet1", exists: true }],
    tables: [],
    ranges: [
      { sheet: "Sheet1", address: "A1", resolvedAddress: "Sheet1!A1", rowCount: 1, columnCount: 1, sentinel: "before", values: [["before"]], numberFormat: "General" },
      { sheet: "Sheet1", address: "B1", resolvedAddress: "Sheet1!B1", rowCount: 1, columnCount: 1, sentinel: null, values: [[null]], numberFormat: "General" },
    ],
    chartSources: [],
    charts: [],
    comments: [],
  };
}

function commentPreconditionManifest(): WorkbookPlanPreconditionManifest {
  return {
    version: 1,
    sheets: [{ name: "Sheet1", exists: true }],
    tables: [],
    ranges: [
      { sheet: "Sheet1", address: "A1", resolvedAddress: "Sheet1!A1", rowCount: 1, columnCount: 1, sentinel: null, values: [[null]], numberFormat: "General" },
      { sheet: "Sheet1", address: "B1", resolvedAddress: "Sheet1!B1", rowCount: 1, columnCount: 1, sentinel: null, values: [[null]], numberFormat: "General" },
    ],
    chartSources: [],
    charts: [],
    comments: [
      { sheet: "Sheet1", address: "B1", exists: false },
      { sheet: "Sheet1", address: "G1", exists: false },
    ],
  };
}

function policyFor(
  effects: readonly Effect[],
  requestClassification: WorkbookPlanPolicyInput["requestClassification"] = "explicit_mutation",
): WorkbookPlanPolicyInput {
  return {
    requestClassification,
    activeSheetName: "Sheet1",
    selectionScopedRequest: false,
    operationSummaries: effects.map((effect, operationIndex) => ({
      operationIndex,
      effect,
      target: { operationIndex, sheetNames: effect === "calculate" ? [] : ["Sheet1"], selectionRelationship: "inside" },
    })),
    capabilitySummaries: ["excel_workbook_access", "workbook_plan_v1_operation_audit"],
    resourceUsage: { operationCount: effects.length, cellsRead: 0, cellsWritten: effects.length, formulaUtf8Bytes: 0, csvPayloadUtf8Bytes: 0, commentUtf8Bytes: 0, sheetsCreated: 0, chartsCreated: 0, tablesCreated: 0, verificationReads: 3 },
    resourceLimits: { operationCount: 100, cellsRead: 100, cellsWritten: 100, formulaUtf8Bytes: 10000, csvPayloadUtf8Bytes: 10000, commentUtf8Bytes: 10000, sheetsCreated: 10, chartsCreated: 10, tablesCreated: 10, verificationReads: 100 },
  };
}

function existingCommentPlan(text: string): Extract<WorkbookPlan, { readonly outcome: "execute" }> {
  return {
    version: 1,
    outcome: "execute",
    summary: "existing comment add-or-update",
    operations: [{ kind: "add_comment", sheet: "Sheet1", address: "B1", text }],
    verification: [{ kind: "comment_exists", sheet: "Sheet1", address: "B1", text }],
  };
}

function deterministicUuidSource(): () => string {
  const values = [
    IDS.turnId,
    "44444444-4444-4444-8444-444444444444",
    "55555555-5555-4555-8555-555555555555",
  ];
  let index = 0;
  return () => values[index++]!;
}

describe("Task 11 Slice C1 real Office.js mutation semantics", () => {
  it("captures an existing comment precondition and updates it through host-backed preparation", async () => {
    const plan = existingCommentPlan("new content");
    const fake = createStrictFake({ initialComments: [{ address: "B1", content: "old content" }] });
    await withExcel(fake, async () => {
      const runtime = await readNativeWorkbookPlanRuntimeContext(plan);
      expect(runtime.compileContext.preconditionManifest?.comments).toEqual([
        { sheet: "Sheet1", address: "B1", exists: true, text: "old content" },
      ]);
      const prepared = await prepareWorkbookPlanExecution({
        plan,
        compileContext: runtime.compileContext,
        policyInput: policyFor(["write"]),
        trustedOperationBoundary: createTrustedOperationBoundary(IDS),
        uuidSource: deterministicUuidSource(),
      });
      expect(prepared.accepted).toBe(true);
      if (!prepared.accepted) return;
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        createTrustedOperationBoundary(IDS),
        createNativeWorkbookPlanAdapter(),
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("committed");
    });
    expect(fake.calls.filter((call) => call.kind === "comment.add")).toHaveLength(0);
    expect(fake.calls.filter((call) => call.kind === "comment.update")).toEqual([
      { kind: "comment.update", args: ["Sheet1", "B1", "new content"] },
    ]);
  });

  it("discards an existing-comment update when Excel.run rejects before hidden sync", async () => {
    const fake = createStrictFake({
      initialComments: [{ address: "B1", content: "old content" }],
      onMutationQueued: () => { throw new Error("forced callback rejection"); },
    });
    await expect(executeCompilerOperations(fake, [
      { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "new content" },
    ])).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });

    const comments = fake.sheet.comments;
    comments.load("items");
    await fake.context.sync();
    expect(comments.items).toHaveLength(1);
    expect(comments.items[0].__contentValue).toBe("old content");
  });

  it("keeps all modeled range and comment mutations pending until explicit or hidden sync", async () => {
    const mutate = (targets: { range: AnyRecord; comment: AnyRecord }): void => {
      targets.range.values = [["after"]];
      targets.range.formulas = [["=1+1"]];
      targets.range.numberFormat = [["0.00"]];
      targets.comment.content = "new content";
    };

    const explicit = createStrictFake({
      initialValues: { "Sheet1!A1": "before" },
      initialNumberFormats: { "Sheet1!A1": "General" },
      initialComments: [{ address: "B1", content: "old content" }],
    });
    const explicitTargets = await loadMutationTargets(explicit);
    mutate(explicitTargets);
    expect(explicitTargets.range.values).toEqual([["before"]]);
    expect(explicitTargets.range.formulas).toEqual([[""]]);
    expect(explicitTargets.range.numberFormat).toEqual([["General"]]);
    expect(explicitTargets.comment.content).toBe("old content");
    expect(explicit.unsynchronizedMutationCommands).toBe(4);
    await explicit.context.sync();
    expect(explicitTargets.range.values).toEqual([["after"]]);
    expect(explicitTargets.range.formulas).toEqual([["=1+1"]]);
    expect(explicitTargets.range.numberFormat).toEqual([["0.00"]]);
    expect(explicitTargets.comment.content).toBe("new content");
    expect(explicit.committedMutationCommands).toBe(4);

    const hidden = createStrictFake({
      initialValues: { "Sheet1!A1": "before" },
      initialNumberFormats: { "Sheet1!A1": "General" },
      initialComments: [{ address: "B1", content: "old content" }],
    });
    const hiddenTargets = await loadMutationTargets(hidden);
    await withExcel(hidden, async () => {
      await (globalThis as AnyRecord).Excel.run(async () => {
        mutate(hiddenTargets);
        expect(hiddenTargets.range.values).toEqual([["before"]]);
        expect(hiddenTargets.range.formulas).toEqual([[""]]);
        expect(hiddenTargets.range.numberFormat).toEqual([["General"]]);
        expect(hiddenTargets.comment.content).toBe("old content");
      });
    });
    expect(hiddenTargets.range.values).toEqual([["after"]]);
    expect(hiddenTargets.range.formulas).toEqual([["=1+1"]]);
    expect(hiddenTargets.range.numberFormat).toEqual([["0.00"]]);
    expect(hiddenTargets.comment.content).toBe("new content");
    expect(hidden.hiddenFinalSyncCount).toBe(1);
    expect(hidden.committedMutationCommands).toBe(4);

    const rejected = createStrictFake({
      initialValues: { "Sheet1!A1": "before" },
      initialNumberFormats: { "Sheet1!A1": "General" },
      initialComments: [{ address: "B1", content: "old content" }],
    });
    const rejectedTargets = await loadMutationTargets(rejected);
    await expect(withExcel(rejected, async () => (globalThis as AnyRecord).Excel.run(async () => {
      mutate(rejectedTargets);
      throw new Error("forced callback rejection");
    }))).rejects.toThrow("forced callback rejection");
    expect(rejectedTargets.range.values).toEqual([["before"]]);
    expect(rejectedTargets.range.formulas).toEqual([[""]]);
    expect(rejectedTargets.range.numberFormat).toEqual([["General"]]);
    expect(rejectedTargets.comment.content).toBe("old content");
    expect(rejected.unsynchronizedMutationCommands).toBe(0);
    expect(rejected.committedMutationCommands).toBe(0);
    expect(rejected.hiddenFinalSyncCount).toBe(0);
  });

  it("fails an existing-comment update as stale after external content changes", async () => {
    const plan = existingCommentPlan("new content");
    const fake = createStrictFake({ initialComments: [{ address: "B1", content: "old content" }] });
    await withExcel(fake, async () => {
      const runtime = await readNativeWorkbookPlanRuntimeContext(plan);
      expect(runtime.compileContext.preconditionManifest?.comments).toEqual([
        { sheet: "Sheet1", address: "B1", exists: true, text: "old content" },
      ]);
      fake.setCommentContent("B1", "external content");
      const prepared = await prepareWorkbookPlanExecution({
        plan,
        compileContext: runtime.compileContext,
        policyInput: policyFor(["write"]),
        trustedOperationBoundary: createTrustedOperationBoundary(IDS),
        uuidSource: deterministicUuidSource(),
      });
      expect(prepared.accepted).toBe(true);
      if (!prepared.accepted) return;
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        createTrustedOperationBoundary(IDS),
        createNativeWorkbookPlanAdapter(),
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("stale");
      expect(report.verification).toBe("not_attempted");
    });
    expect(fake.calls.filter((call) => ["comment.add", "comment.update", "values", "formulas", "numberFormat", "calculate"].includes(call.kind))).toHaveLength(0);
  });

  it("fails a later-group add_comment when cancellation arrives during comment pre-read", async () => {
    let releaseCommentScan!: () => void;
    const pendingCommentScan = new Promise<void>((resolve) => { releaseCommentScan = resolve; });
    const fake = createStrictFake({ pauseCommentScanSync: { pending: pendingCommentScan } });
    const controller = new AbortController();
    const adapter = createNativeWorkbookPlanAdapter();
    const group = compilerGroupForOperations([
      { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "must not add" },
    ]);
    const execution = withExcel(fake, () => adapter.executeGroup(group, 1, controller.signal));
    for (let attempt = 0; attempt < 20 && !fake.commentScanSyncStarted; attempt += 1) {
      await new Promise<void>((resolve) => queueMicrotask(resolve));
    }
    expect(fake.commentScanSyncStarted).toBe(true);
    controller.abort();
    releaseCommentScan();
    await expect(execution).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    expect(fake.calls.filter((call) => ["comment.add", "comment.update", "values", "formulas", "numberFormat", "calculate"].includes(call.kind))).toHaveLength(0);
  });

  it("rejects the real first-group Excel.run gate before hidden final sync and reconciles honestly", async () => {
    const controller = new AbortController();
    const fake = createStrictFake({ onMutationQueued: () => controller.abort() });
    const baseAdapter = createNativeWorkbookPlanAdapter();
    let laterGroupCalls = 0;
    let verificationCalls = 0;
    const adapter = {
      ...baseAdapter,
      executeGroup: async (...args: Parameters<typeof baseAdapter.executeGroup>) => {
        laterGroupCalls += 1;
        return baseAdapter.executeGroup(...args);
      },
      verifyPlan: async (...args: Parameters<typeof baseAdapter.verifyPlan>) => {
        verificationCalls += 1;
        return baseAdapter.verifyPlan(...args);
      },
    };
    const plan: WorkbookPlan = {
      version: 1,
      outcome: "execute",
      summary: "real first-group cancellation gate",
      operations: [
        { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "must not commit" },
        { kind: "calculate" },
      ],
      verification: [{ kind: "calculation_completed" }],
    };
    const prepared = await prepareWorkbookPlanExecution({
      plan,
      compileContext: compileContext(commentPreconditionManifest()),
      policyInput: policyFor(["write", "calculate"]),
      trustedOperationBoundary: createTrustedOperationBoundary(IDS),
      uuidSource: deterministicUuidSource(),
    });
    expect(prepared.accepted).toBe(true);
    if (!prepared.accepted) return;

    await withExcel(fake, async () => {
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        createTrustedOperationBoundary(IDS),
        adapter,
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
        controller.signal,
      );
      expect(report.outcome).toBe("partial");
      expect(report.firstFailedGroup).toBe(0);
      expect(report.verification).toBe("reconciled");
      expect(report.rollback).toBe("not_attempted");
      expect(report.operations[0]?.state).toBe("not_committed");
      expect(report.operations[0]?.status).toBe("failed");
      expect(report.operations[1]?.status).toBe("not_started");
    });

    expect(laterGroupCalls).toBe(0);
    expect(verificationCalls).toBe(0);
    expect(fake.hiddenFinalSyncCount).toBe(0);
    expect(fake.committedMutationCommands).toBe(0);
    expect(fake.unsynchronizedMutationCommands).toBe(0);
    expect(fake.calls.filter((call) => call.kind === "comment.add")).toHaveLength(1);
  });

  it("skips Office hidden final sync after later-group cancellation", async () => {
    const controller = new AbortController();
    const fake = createStrictFake({
      initialComments: [{ address: "B1", content: "old" }],
      onMutationQueued: () => controller.abort(),
    });
    const adapter = createNativeWorkbookPlanAdapter();
    const group = compilerGroupForOperations([
      { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "new" },
    ]);
    const result = await withExcel(fake, () => adapter.executeGroup(group, 1, controller.signal));
    expect(result).toEqual({ status: "sync_failed", failureCode: "sync_failed" });
    expect(fake.hiddenFinalSyncCount).toBe(0);
    expect(fake.committedMutationCommands).toBe(0);
    expect(fake.unsynchronizedMutationCommands).toBe(0);
    expect(fake.calls.filter((call) => call.kind === "sync")).toHaveLength(3);
  });

  it("fails an over-limit add_comment before items, location, content, or mutation work", async () => {
    const fake = createStrictFake({ reportedCommentCount: 513 });
    await expect(executeCompilerOperations(fake, [
      { kind: "add_comment", sheet: "Sheet1", address: "A1", text: "blocked" },
    ])).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });

    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => [
      "comments.load.items", "comment.getLocation", "comment.content.load", "comment.add", "comment.update",
    ].includes(call.kind))).toHaveLength(0);
    expect(fake.runCount).toBe(1);
  });

  it("fails comment_exists verification on an over-limit collection before item or per-comment scans", async () => {
    const fake = createStrictFake({ reportedCommentCount: 513 });
    const adapter = createNativeWorkbookPlanAdapter();
    await withExcel(fake, async () => {
      const report = await adapter.verifyPlan([
        { kind: "comment_exists", sheet: "Sheet1", address: "A1", text: "blocked" }
      ]);
      expect(report.ok).toBe(false);
      expect(report.counts.error).toBe(1);
    });

    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => [
      "comments.load.items", "comment.getLocation", "comment.content.load",
    ].includes(call.kind))).toHaveLength(0);
  });

  it("fails the first-group comment precondition conservatively without mutation, later groups, or verification", async () => {
    const plan: WorkbookPlan = {
      version: 1,
      outcome: "execute",
      summary: "bounded first-group comment precondition",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [["must not write"]] },
        { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "must not add" },
        { kind: "calculate" },
      ],
      verification: [{ kind: "calculation_completed" }],
    };
    const prepared = await prepareWorkbookPlanExecution({
      plan,
      compileContext: compileContext(commentPreconditionManifest()),
      policyInput: policyFor(["write", "write", "calculate"]),
      trustedOperationBoundary: createTrustedOperationBoundary(IDS),
      uuidSource: (() => {
        const values = [IDS.turnId, "44444444-4444-4444-8444-444444444444", "55555555-5555-4555-8555-555555555555", "66666666-6666-4666-8666-666666666666"];
        let index = 0;
        return () => values[index++]!;
      })(),
    });
    expect(prepared.accepted).toBe(true);
    if (!prepared.accepted) return;
    const fake = createStrictFake({ reportedCommentCount: 513 });
    const adapter = createNativeWorkbookPlanAdapter();
    let verificationCalls = 0;
    const executionAdapter = {
      ...adapter,
      verifyPlan: async (...args: Parameters<typeof adapter.verifyPlan>) => {
        verificationCalls += 1;
        return adapter.verifyPlan(...args);
      },
    };
    await withExcel(fake, async () => {
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        createTrustedOperationBoundary(IDS),
        executionAdapter,
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("partial");
      expect(report.firstFailedGroup).toBe(0);
      expect(report.verification).toBe("reconciled");
      expect(report.operations[0]?.status).toBe("unknown_within_batch");
      expect(report.operations[1]?.status).toBe("unknown_within_batch");
      expect(report.operations[2]?.status).toBe("not_started");
    });

    expect(verificationCalls).toBe(0);
    expect(fake.calls.filter((call) => [
      "values", "comment.add", "comment.update", "calculate", "comments.load.items", "comment.getLocation", "comment.content.load",
    ].includes(call.kind))).toHaveLength(0);
  });

  it("uses one logical-sheet comment scan for a mandatory first-group precondition and add_comment", async () => {
    const fake = createStrictFake({ freshWorksheetProxies: true });
    const adapter = createNativeWorkbookPlanAdapter();
    const group = compilerGroupForOperations([
      { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "created" },
    ]);
    const gate: WorkbookPlanFirstGroupContext = {
      preconditionManifest: commentPreconditionManifest(),
      executeAfterFresh: async (_observation, mutation) => mutation(group),
    };

    await withExcel(fake, async () => {
      await expect(adapter.executeFirstGroup(gate)).resolves.toEqual({ status: "committed" });
    });

    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comments.load.items")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comment.add")).toEqual([
      { kind: "comment.add", args: ["Sheet1!B1", "created"] },
    ]);
  });

  it("scans duplicate same-sheet verification requests once across distinct worksheet proxies", async () => {
    const fake = createStrictFake({
      freshWorksheetProxies: true,
      initialComments: [
        { address: "B1", content: "first" },
        { address: "G1", content: "second" },
      ],
    });
    const adapter = createNativeWorkbookPlanAdapter();

    await withExcel(fake, async () => {
      const report = await adapter.verifyPlan([
        { kind: "comment_exists", sheet: "Sheet1", address: "B1", text: "first" },
        { kind: "comment_exists", sheet: "Sheet1", address: "G1", text: "second" },
      ]);
      expect(report.ok).toBe(true);
      expect(report.counts.verified).toBe(2);
    });

    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comments.load.items")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comment.getLocation")).toHaveLength(2);
    expect(fake.calls.filter((call) => call.kind === "comment.content.load")).toHaveLength(2);
  });

  it("scans duplicate same-sheet reconciliation requests once across distinct worksheet proxies", async () => {
    const fake = createStrictFake({
      freshWorksheetProxies: true,
      initialComments: [
        { address: "B1", content: "first" },
        { address: "G1", content: "second" },
      ],
    });
    const adapter = createNativeWorkbookPlanAdapter();
    const group = compilerGroupForOperations([
      { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "first" },
      { kind: "add_comment", sheet: "Sheet1", address: "G1", text: "second" },
    ]);

    await withExcel(fake, async () => {
      await expect(adapter.reconcileFailedGroup(group)).resolves.toEqual(["committed", "committed"]);
    });

    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comments.load.items")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comment.getLocation")).toHaveLength(2);
    expect(fake.calls.filter((call) => call.kind === "comment.content.load")).toHaveLength(2);
  });

  it("scans one worksheet collection once for multiple add_comment operations and preserves same-cell sequencing", async () => {
    const fake = createStrictFake({ initialComments: [{ address: "E1", content: "old" }] });
    await expect(executeCompilerOperations(fake, [
      { kind: "add_comment", sheet: "Sheet1", address: "E1", text: "updated" },
      { kind: "add_comment", sheet: "Sheet1", address: "F1", text: "new" },
      { kind: "add_comment", sheet: "Sheet1", address: "E1", text: "updated again" },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comments.load.items")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comment.getLocation")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comment.content.load")).toHaveLength(1);
    expect(fake.calls.filter((call) => call.kind === "comment.add")).toEqual([
      { kind: "comment.add", args: ["Sheet1!F1", "new"] },
    ]);
    expect(fake.calls.filter((call) => call.kind === "comment.update")).toEqual([
      { kind: "comment.update", args: ["Sheet1", "E1", "updated"] },
      { kind: "comment.update", args: ["Sheet1", "E1", "updated again"] },
    ]);
    expect(fake.runCount).toBe(1);
  });

  it("performs the comment pre-read before any earlier non-comment mutation in its group", async () => {
    const fake = createStrictFake({ reportedCommentCount: 513 });
    await expect(executeCompilerOperations(fake, [
      { kind: "write_values", sheet: "Sheet1", address: "A1", values: [["must not write"]] },
      { kind: "add_comment", sheet: "Sheet1", address: "B1", text: "blocked" },
    ])).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });

    expect(fake.calls.filter((call) => ["values", "comment.add", "comment.update"].includes(call.kind))).toHaveLength(0);
    expect(fake.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(fake.calls.filter((call) => ["comments.load.items", "comment.getLocation", "comment.content.load"].includes(call.kind))).toHaveLength(0);
  });

  it("accepts the 512-item boundary and fails closed for malformed comment counts", async () => {
    const boundary = createStrictFake({ reportedCommentCount: 512, syntheticCommentCount: 512 });
    await expect(executeCompilerOperations(boundary, [
      { kind: "add_comment", sheet: "Sheet1", address: "A1", text: "boundary" },
    ])).resolves.toEqual({ status: "committed" });
    expect(boundary.calls.filter((call) => call.kind === "comments.getCount")).toHaveLength(1);
    expect(boundary.calls.filter((call) => call.kind === "comments.load.items")).toHaveLength(1);
    expect(boundary.calls.filter((call) => call.kind === "comment.getLocation")).toHaveLength(512);
    expect(boundary.calls.filter((call) => call.kind === "comment.content.load")).toHaveLength(512);

    for (const reportedCommentCount of [-1, 1.5, Number.NaN, "512", null]) {
      const malformed = createStrictFake({ reportedCommentCount });
      await expect(executeCompilerOperations(malformed, [
        { kind: "add_comment", sheet: "Sheet1", address: "A1", text: "blocked" },
      ])).resolves.toEqual({ status: "sync_failed", failureCode: "sync_failed" });
      expect(malformed.calls.filter((call) => [
        "comments.load.items", "comment.getLocation", "comment.content.load", "comment.add", "comment.update",
      ].includes(call.kind))).toHaveLength(0);
    }
  });

  it("clears a compiler-emitted table filter through the explicit reset operation", async () => {
    const fake = createStrictFake();
    await expect(executeCompilerOperations(fake, [
      { kind: "reset_table_filter", tableName: "Table1", columnIndex: 0 },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "filter.clear")).toEqual([
      { kind: "filter.clear", args: ["Table1", 0] },
    ]);
    expect(fake.calls.filter((call) => call.kind === "filter.custom")).toHaveLength(0);
  });

  it("passes only criterion1 for a compiler-emitted one-criterion custom filter", async () => {
    const fake = createStrictFake();
    await expect(executeCompilerOperations(fake, [
      { kind: "filter_table", tableName: "Table1", columnIndex: 0, criterion1: ">0" },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "filter.custom")).toEqual([
      { kind: "filter.custom", args: ["Table1", 0, ">0"] },
    ]);
  });

  it("passes criterion1 and criterion2 without inventing an operator", async () => {
    const fake = createStrictFake();
    await expect(executeCompilerOperations(fake, [
      { kind: "filter_table", tableName: "Table1", columnIndex: 0, criterion1: ">0", criterion2: "<10" },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "filter.custom")).toEqual([
      { kind: "filter.custom", args: ["Table1", 0, ">0", "<10"] },
    ]);
  });

  it("preserves the optional operator position when criterion2 is omitted", async () => {
    const fake = createStrictFake();
    await expect(executeCompilerOperations(fake, [
      { kind: "filter_table", tableName: "Table1", columnIndex: 0, criterion1: ">0", operator: "Or" },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "filter.custom")).toEqual([
      { kind: "filter.custom", args: ["Table1", 0, ">0", undefined, "Or"] },
    ]);
  });

  it("defaults omitted compiler-emitted conditional operators to the documented Office values", async () => {
    const fake = createStrictFake();
    await expect(executeCompilerOperations(fake, [
      {
        kind: "conditional_format",
        sheet: "Sheet1",
        address: "A1:C2",
        ruleType: "cellValue",
        formula1: "0",
        fillColor: "#00ff00",
        fontColor: "#000000",
      },
      {
        kind: "conditional_format",
        sheet: "Sheet1",
        address: "A1:C2",
        ruleType: "containsText",
        text: "ok",
        fillColor: "#00ff00",
        fontColor: "#000000",
      },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "CellValue.rule")).toEqual([
      { kind: "CellValue.rule", args: [{ operator: "GreaterThan", formula1: "0" }] },
    ]);
    expect(fake.calls.filter((call) => call.kind === "ContainsText.rule")).toEqual([
      { kind: "ContainsText.rule", args: [{ operator: "Contains", text: "ok" }] },
    ]);
  });

  it("emits complete compiler-emitted validation prompts for either partial prompt field", async () => {
    const fake = createStrictFake();
    await expect(executeCompilerOperations(fake, [
      { kind: "set_validation_list", sheet: "Sheet1", address: "A1", source: "one,two", promptTitle: "Pick" },
      { kind: "set_validation_list", sheet: "Sheet1", address: "A2", source: "one,two", promptMessage: "Choose one" },
      { kind: "set_validation_list", sheet: "Sheet1", address: "A3", source: "one,two" },
    ])).resolves.toEqual({ status: "committed" });

    expect(fake.calls.filter((call) => call.kind === "validation.prompt")).toEqual([
      { kind: "validation.prompt", args: [{ showPrompt: true, title: "Pick", message: "" }] },
      { kind: "validation.prompt", args: [{ showPrompt: true, title: "", message: "Choose one" }] },
    ]);
  });

  it("executes every compiler-produced operation kind through strict API-shaped Office fakes", async () => {
    const compiled = compilerGroups(semanticPlan());
    expect(compiled.primitives.map((primitive) => primitive.kind)).toEqual([
      "write_values", "write_formula", "fill_formula", "fill_formula", "add_comment", "add_comment",
      "format_range", "set_validation_list", "conditional_format", "conditional_format", "freeze_panes", "freeze_panes",
      "create_table", "create_chart", "create_sheet", "rename_sheet", "copy_sheet", "sort_table", "filter_table", "filter_table", "reset_table_filter", "calculate",
    ]);
    const fake = createStrictFake();
    const adapter = createNativeWorkbookPlanAdapter();
    await withExcel(fake, async () => {
      for (const group of compiled.groups) {
        await expect(adapter.executeGroup(group, group.groupIndex)).resolves.toEqual({ status: "committed" });
      }
    });

    const autoFills = fake.calls.filter((call) => call.kind === "autoFill");
    expect(autoFills).toHaveLength(2);
    expect(autoFills[0].args).toEqual(["Sheet1!D1", "Sheet1!D1:F1"]);
    expect(autoFills[1].args).toEqual(["Sheet1!D1", "Sheet1!D1:D3"]);
    expect(fake.calls.find((call) => call.kind === "values")?.args).toEqual(["Sheet1!A1:C2", [[1, 2, 3], [4, 5, 6]]]);
    expect(fake.calls.find((call) => call.kind === "formulas")?.args).toEqual(["Sheet1!D1", [["=A1+B1"]]]);
    expect(fake.calls.some((call) => call.kind === "numberFormat" && call.args[0] === "Sheet1!A1:C2" && JSON.stringify(call.args[1]) === JSON.stringify([["0.00", "0.00", "0.00"], ["0.00", "0.00", "0.00"]]))).toBe(true);
    for (const [kind, expected] of [
      ["font.bold", true], ["font.italic", true], ["font.color", "#111111"], ["font.name", "Aptos"], ["font.size", 12], ["font.underline", "Single"],
      ["fill.color", "#eeeeee"], ["format.horizontalAlignment", "Center"], ["format.verticalAlignment", "Center"], ["format.wrapText", true], ["format.indentLevel", 2], ["format.textOrientation", 15],
    ] as const) {
      expect(fake.calls.some((call) => call.kind === kind && call.args[1] === expected)).toBe(true);
    }
    for (const border of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"]) {
      expect(fake.calls.some((call) => call.kind === "border.style" && call.args[1] === border && call.args[2] === "Continuous")).toBe(true);
      expect(fake.calls.some((call) => call.kind === "border.color" && call.args[1] === border && call.args[2] === "#222222")).toBe(true);
      expect(fake.calls.some((call) => call.kind === "border.weight" && call.args[1] === border && call.args[2] === "Medium")).toBe(true);
    }
    expect(fake.calls.some((call) => call.kind === "autofitColumns" && call.args[0] === "Sheet1!A1:C2")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "autofitRows" && call.args[0] === "Sheet1!A1:C2")).toBe(true);
    expect(fake.calls.filter((call) => call.kind === "table.add")).toEqual([{ kind: "table.add", args: ["Sheet1!A1:C4", true] }]);
    expect(fake.calls.filter((call) => call.kind === "table.sort")[0]?.args[0]).toEqual([{ key: 1, ascending: false }]);
    expect(fake.calls.some((call) => call.kind === "filter.values" && call.args[1] === 0 && JSON.stringify(call.args[2]) === JSON.stringify(["one", "two"]))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "filter.custom" && JSON.stringify(call.args.slice(1)) === JSON.stringify([1, ">0", "<10", "And"]))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "filter.clear" && call.args[1] === 0)).toBe(true);
    expect(fake.calls.some((call) => call.kind === "chart.add" && JSON.stringify(call.args.slice(0, 3)) === JSON.stringify(["ColumnClustered", "Sheet1!A1:C4", "Auto"]))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "chart.title" && call.args[0] === "Revenue")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "validation.rule" && JSON.stringify(call.args[0]) === JSON.stringify({ list: { inCellDropDown: true, source: "one,two" } }))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "validation.prompt" && JSON.stringify(call.args[0]) === JSON.stringify({ showPrompt: true, title: "Pick", message: "Choose one" }))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "CellValue.rule" && JSON.stringify(call.args[0]) === JSON.stringify({ operator: "GreaterThan", formula1: "0" }))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "CellValue.format.fill" && call.args[0] === "#00ff00")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "CellValue.format.font" && call.args[0] === "#000000")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "ContainsText.rule" && JSON.stringify(call.args[0]) === JSON.stringify({ operator: "Contains", text: "ok" }))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "ContainsText.format.fill" && call.args[0] === "#00ff00")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "ContainsText.format.font" && call.args[0] === "#000000")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "freezeRows" && call.args[0] === 2)).toBe(true);
    expect(fake.calls.some((call) => call.kind === "freezeColumns" && call.args[0] === 3)).toBe(true);
    expect(fake.calls.some((call) => call.kind === "unfreeze")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "sheet.add" && call.args[0] === "NewSheet")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "sheet.rename" && call.args[0] === "NewSheet" && call.args[1] === "RenamedSheet")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "sheet.copy" && call.args[0] === "End")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "sheet.rename" && call.args[1] === "CopiedSheet")).toBe(true);
    expect(fake.calls.filter((call) => call.kind === "comment.add")).toHaveLength(1);
    expect(fake.calls.some((call) => call.kind === "comment.add" && JSON.stringify(call.args) === JSON.stringify(["Sheet1!E1", "first"]))).toBe(true);
    expect(fake.calls.some((call) => call.kind === "comment.update" && call.args[0] === "Sheet1" && call.args[1] === "E1" && call.args[2] === "updated")).toBe(true);
    expect(fake.calls.some((call) => call.kind === "calculate" && call.args[0] === "Full")).toBe(true);
    expect(fake.runCount).toBe(compiled.groups.length);
  });

  it("R8 production seam fails closed for unavailable validation and conditional-format owners", async () => {
    const runUnavailable = async (
      unavailable: readonly ("validation" | "conditional")[],
      operation: WorkbookOperation,
      effect: Effect,
    ): Promise<void> => {
      const fake = createStrictFake({ unavailable });
      const plan: Extract<WorkbookPlan, { readonly outcome: "execute" }> = {
        version: 1,
        outcome: "execute",
        summary: `R8 unavailable ${operation.kind}`,
        operations: [operation],
        verification: [{ kind: "calculation_completed" }],
      };
      await withExcel(fake, async () => {
        const runtime = await readNativeWorkbookPlanRuntimeContext(plan);
        const prepared = await prepareWorkbookPlanExecution({
          plan,
          compileContext: runtime.compileContext,
          policyInput: policyFor([effect], "explicit_structure"),
          trustedOperationBoundary: createTrustedOperationBoundary(IDS),
          uuidSource: deterministicUuidSource(),
        });
        expect(prepared.accepted).toBe(true);
        if (!prepared.accepted) return;
        const report = await executeAcceptedWorkbookPlan(
          prepared.transaction,
          createTrustedOperationBoundary(IDS),
          createNativeWorkbookPlanAdapter(),
          new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
        );
        expect(report.outcome).not.toBe("committed");
        expect(report.verification).toBe("reconciled");
      });
    };
    await runUnavailable(
      ["validation"],
      { kind: "set_validation_list", sheet: "Sheet1", address: "F1:F2", source: "one,two" },
      "write",
    );
    await runUnavailable(
      ["conditional"],
      { kind: "conditional_format", sheet: "Sheet1", address: "A1:C2", ruleType: "cellValue", formula1: "0", fillColor: "#00ff00", fontColor: "#000000" },
      "format",
    );
  });

  it("preserves load-then-sync lifecycle and uses bounded comment location selection for reads", async () => {
    const fake = createStrictFake({ initialComments: [{ address: "G1", content: "existing" }] });
    const range = fake.sheet.getRange("A1");
    expect(() => range.values).toThrow(/before load\+sync/);
    const adapter = createNativeWorkbookPlanAdapter();
    await withExcel(fake, async () => {
      const report = await adapter.verifyPlan([
        { kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [[null]] },
        { kind: "comment_exists", sheet: "Sheet1", address: "G1", text: "existing" },
        { kind: "calculation_completed" },
      ]);
      expect(report.ok).toBe(true);
      expect(report.counts.verified).toBe(3);
    });
    expect(fake.syncCount).toBeGreaterThanOrEqual(2);
  });

  it("maps an operation-specific Office failure to conservative group failure without later mutation or verification", async () => {
    const plan: WorkbookPlan = {
      version: 1,
      outcome: "execute",
      summary: "C1 error seam",
      operations: [
        { kind: "write_values", sheet: "Sheet1", address: "A1", values: [["after"]] },
        { kind: "format_range", sheet: "Sheet1", address: "B1", format: { numberFormat: "0.00" } },
        { kind: "calculate" },
      ],
      verification: [
        { kind: "range_values_equal", sheet: "Sheet1", address: "A1", expectedValues: [["after"]] },
        { kind: "range_number_format_equal", sheet: "Sheet1", address: "B1", expectedNumberFormat: "0.00" },
        { kind: "calculation_completed" },
      ],
    };
    const prepared = await prepareWorkbookPlanExecution({
      plan,
      compileContext: compileContext(preconditionManifest()),
      policyInput: policyFor(["write", "format", "calculate"]),
      trustedOperationBoundary: createTrustedOperationBoundary(IDS),
      uuidSource: (() => {
        const values = [IDS.turnId, "44444444-4444-4444-8444-444444444444", "55555555-5555-4555-8555-555555555555", "66666666-6666-4666-8666-666666666666"];
        let index = 0;
        return () => values[index++]!;
      })(),
    });
    expect(prepared.accepted).toBe(true);
    if (!prepared.accepted) return;
    const fake = createStrictFake({
      failOnNumberFormat: true,
      initialValues: { "Sheet1!A1": "before", "Sheet1!B1": null },
      initialNumberFormats: { "Sheet1!A1": "General", "Sheet1!B1": "General" },
    });
    const adapter = createNativeWorkbookPlanAdapter();
    let verificationCalls = 0;
    const executionAdapter = {
      ...adapter,
      verifyPlan: async (...args: Parameters<typeof adapter.verifyPlan>) => {
        verificationCalls += 1;
        return adapter.verifyPlan(...args);
      },
    };
    await withExcel(fake, async () => {
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        createTrustedOperationBoundary(IDS),
        executionAdapter,
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("partial");
      expect(report.firstFailedGroup).toBe(1);
      expect(report.verification).toBe("reconciled");
      expect(verificationCalls).toBe(0);
      expect(fake.calls.some((call) => call.kind === "calculate")).toBe(false);
      expect(fake.calls.some((call) => ["values", "formulas", "numberFormat", "table.add", "table.sort", "filter.values", "filter.custom", "filter.clear", "chart.add", "validation.rule", "conditional.add", "freezeRows", "freezeColumns", "unfreeze", "sheet.add", "sheet.rename", "sheet.copy", "comment.add", "comment.update"].includes(call.kind))).toBe(true);
      expect(fake.runCount).toBe(3);
    });
  });

  it("R8 executes operation-class plans through preparation, native groups, mandatory proofs, and final reports", async () => {
    const fake = createStrictFake();
    const trustedBoundary = createTrustedOperationBoundary(IDS);
    const uuidValues = [IDS.turnId, ...Array.from({ length: 64 }, (_, index) => `00000000-0000-4000-8000-${String(index + 1).padStart(12, "0")}`)];
    let uuidIndex = 0;
    const runSeam = async (operations: readonly WorkbookOperation[], effect: Effect): Promise<void> => {
      const plan: Extract<WorkbookPlan, { readonly outcome: "execute" }> = {
        version: 1,
        outcome: "execute",
        summary: `R8 ${operations[0]?.kind ?? "operation"}`,
        operations,
        verification: [{ kind: "calculation_completed" }],
      };
      const runtime = await readNativeWorkbookPlanRuntimeContext(plan);
      const compileContext: VerifiedWorkbookCompileContext = {
        ...runtime.compileContext,
        sheetNames: new Set([
          ...runtime.compileContext.sheetNames,
          ...operations.flatMap((operation) => operation.kind === "rename_sheet"
            ? [operation.oldName]
            : operation.kind === "copy_sheet"
              ? [operation.sourceName]
              : []),
        ]),
        tablesByName: new Map([["Table1", { canonicalName: "Table1", columnCount: 3, rowCount: 4, sheetName: "Sheet1" }]]),
        chartSourceResolutionsBySheet: new Map([["Sheet1", new Map([["A1:C4", { areaCount: 1 }]])]]),
      };
      const prepared = await prepareWorkbookPlanExecution({
        plan,
        compileContext,
        policyInput: policyFor(Array.from({ length: operations.length }, () => effect), "explicit_structure"),
        trustedOperationBoundary: trustedBoundary,
        uuidSource: () => uuidValues[uuidIndex++]!,
      });
      expect(prepared.accepted).toBe(true);
      if (!prepared.accepted) return;
      const baseAdapter = createNativeWorkbookPlanAdapter();
      let mandatoryVerificationCalls = 0;
      const executionAdapter = {
        ...baseAdapter,
        verifyMandatoryPostconditions: async (...args: Parameters<NonNullable<typeof baseAdapter.verifyMandatoryPostconditions>>) => {
          mandatoryVerificationCalls += 1;
          return baseAdapter.verifyMandatoryPostconditions!(...args);
        },
      };
      const report = await executeAcceptedWorkbookPlan(
        prepared.transaction,
        trustedBoundary,
        executionAdapter,
        new BrowserIdempotencyStore(new MemoryIdempotencyStorage()),
      );
      expect(report.outcome).toBe("committed");
      expect(mandatoryVerificationCalls).toBe(1);
      expect(report.verification).toBe("not_attempted");
      expect(report.operationCount).toBe(operations.length);
    };

    await withExcel(fake, async () => {
      await runSeam([{ kind: "write_values", sheet: "Sheet1", address: "A1:C2", values: [[1, 2, 3], [4, 5, 6]] }], "write");
      await runSeam([{ kind: "write_formula", sheet: "Sheet1", address: "D1", formula: "=A1+B1" }], "write");
      await runSeam([{ kind: "fill_formula", sheet: "Sheet1", sourceAddress: "D1", targetAddress: "D1:F1" }], "write");
      await runSeam([{ kind: "add_comment", sheet: "Sheet1", address: "E1", text: "R8 comment" }], "write");
      await runSeam([{ kind: "format_range", sheet: "Sheet1", address: "A1", format: { numberFormat: "0.00" } }], "format");
      await runSeam([{
        kind: "format_range", sheet: "Sheet1", address: "A1:C2", format: {
          bold: true, italic: true, fontColor: "#111111", fillColor: "#eeeeee", numberFormat: "0.00", borders: true,
          borderColor: "#222222", borderWeight: "Medium", columnAutofit: true, rowAutofit: true,
          horizontalAlignment: "Center", verticalAlignment: "Center", fontName: "Aptos", fontSize: 12,
          underline: "Single", wrapText: true, indentLevel: 2, textOrientation: 15,
        },
      }], "format");
      await runSeam([{ kind: "set_validation_list", sheet: "Sheet1", address: "F1:F2", source: "one,two", promptTitle: "Pick", promptMessage: "Choose one" }], "write");
      await runSeam([{ kind: "conditional_format", sheet: "Sheet1", address: "A1:C2", ruleType: "cellValue", formula1: "0", fillColor: "#00ff00", fontColor: "#000000" }], "format");
      await runSeam([{ kind: "freeze_panes", sheet: "Sheet1", rows: 2, columns: 3 }], "format");
      await runSeam([{ kind: "create_table", sheet: "Sheet1", address: "A1:C4", hasHeaders: true }], "structure");
      await runSeam([{ kind: "sort_table", tableName: "Table1", columnIndex: 1, ascending: false }], "write");
      await runSeam([{ kind: "filter_table", tableName: "Table1", columnIndex: 1, criterion1: ">0", criterion2: "<10", operator: "And" }], "filter");
      await runSeam([{ kind: "reset_table_filter", tableName: "Table1", columnIndex: 0 }], "filter");
      await runSeam([{ kind: "create_chart", sheet: "Sheet1", dataRange: "A1:C4", chartType: "ColumnClustered", title: "Revenue" }], "structure");
      await runSeam([{ kind: "calculate" }], "calculate");
      await runSeam([{ kind: "create_sheet", name: "NewSheet" }], "structure");
      await runSeam([{ kind: "rename_sheet", oldName: "NewSheet", newName: "RenamedSheet" }], "structure");
      await runSeam([{ kind: "copy_sheet", sourceName: "RenamedSheet", newName: "CopiedSheet" }], "structure");
    });
    expect(fake.sortBodyReadCount).toBeGreaterThanOrEqual(2);
  });
});
