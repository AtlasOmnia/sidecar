import type { WorkbookPlanPolicyInput } from "./workbookPlanPolicy";
import {
  createWorkbookPlanPolicyInputSnapshot,
  decideWorkbookPlanPolicy,
  type WorkbookPlanPolicyReasonCode,
} from "./workbookPlanPolicy";
import {
  canonicalizeWorkbookPlan,
  type CanonicalizeWorkbookPlanSuccess,
} from "./workbookPlanCanonicalization";
import {
  compileWorkbookPlan,
  type CompiledWorkbookPrimitive,
  type CompiledWorkbookMandatoryPostcondition,
  type VerifiedWorkbookCompileContext,
  type WorkbookPlanPreconditionManifest,
} from "./workbookPlanCompiler";
import type { WorkbookPlan, VerificationRule } from "./generated/workbookPlanSchema";
import type { WorkbookPlanOperationEffect } from "./workbookPlanPolicy";

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const MAX_BOUNDARY_LENGTH = 256;
const MAX_DIGEST_LENGTH = 128;


export type WorkbookPlanOperationState = "committed" | "not_committed" | "unknown_within_batch";
export type WorkbookPlanOperationStatus = "not_started" | "applied" | "verified" | "failed" | "rolled_back" | "unknown_within_batch";
export type WorkbookPlanExecutionOutcome =
  | "committed"
  | "partial"
  | "stale"
  | "cancelled"
  | "replay"
  | "conflict"
  | "restart_rejected"
  | "rejected"
  | "persistence_unavailable"
  | "persistence_unavailable_after_mutation";
export type WorkbookPlanVerificationStatus = "not_attempted" | "reconciled" | "reconciliation_failed";
export type WorkbookPlanRollbackStatus = "not_attempted" | "unavailable" | "succeeded" | "failed";

export interface WorkbookPlanOperationReport {
  readonly operationId: string;
  readonly operationIndex: number;
  readonly effect: WorkbookPlanOperationEffect;
  readonly groupIndex: number;
  readonly state: WorkbookPlanOperationState;
  readonly status: WorkbookPlanOperationStatus;
}

export interface WorkbookPlanExecutionReport {
  readonly outcome: WorkbookPlanExecutionOutcome;
  readonly transactionId: string;
  readonly operationCount: number;
  readonly operations: readonly WorkbookPlanOperationReport[];
  readonly firstFailedGroup?: number;
  readonly firstFailedOperationIndex?: number;
  readonly verification: WorkbookPlanVerificationStatus;
  readonly rollback: WorkbookPlanRollbackStatus;
}

export interface AcceptedWorkbookPlanTransaction {
  readonly transactionId: string;
  readonly semanticDigest: string;
  readonly operationCount: number;
}

export type WorkbookPlanUuidSource = () => string;
export type WorkbookPlanPreconditionResult = "fresh" | "stale";

export interface TrustedOperationBoundary {
  readonly __trustedOperationBoundary: unique symbol;
}

interface TrustedBoundaryMetadata {
  readonly token: string;
}

const trustedBoundaries = new WeakMap<object, TrustedBoundaryMetadata>();

export function createTrustedOperationBoundary(input: {
  readonly sessionId: string;
  readonly documentId: string;
  readonly turnId: string;
}): TrustedOperationBoundary {
  if (!isUuid(input.sessionId) || !isUuid(input.documentId) || !isUuid(input.turnId)) {
    throw new Error("Trusted operation boundary components must be UUIDs");
  }
  const boundary = Object.freeze({}) as TrustedOperationBoundary;
  trustedBoundaries.set(boundary, Object.freeze({
    token: `${input.sessionId}.${input.documentId}.${input.turnId}`,
  }));
  return boundary;
}

export interface WorkbookPlanInverse {
  readonly completeCapturedPreState: boolean;
  readonly independentlyVerifiedSafe: boolean;
  readonly rollback: () => Promise<boolean>;
}

export interface WorkbookPlanPreparationInput {
  readonly plan: unknown;
  readonly trustedOperationBoundary: TrustedOperationBoundary;
  readonly compileContext: VerifiedWorkbookCompileContext;
  readonly policyInput: WorkbookPlanPolicyInput;
  readonly uuidSource?: WorkbookPlanUuidSource;
  readonly preconditionCheck?: () => Promise<WorkbookPlanPreconditionResult>;
  readonly inverse?: WorkbookPlanInverse;
}

export type WorkbookPlanPreparationResult =
  | { readonly accepted: true; readonly transaction: AcceptedWorkbookPlanTransaction }
  | {
      readonly accepted: false;
      readonly reason: "canonicalization_failed" | "policy_rejected" | "compile_failed" | "invalid_boundary" | "invalid_uuid_source";
      readonly policyReason?: WorkbookPlanPolicyReasonCode;
    };

export interface WorkbookPlanEffectOperation {
  readonly operationId: string;
  readonly operationIndex: number;
  readonly kind: CompiledWorkbookPrimitive["kind"];
  readonly primitiveName: string;
  readonly args: Readonly<Record<string, unknown>>;
}

export interface WorkbookPlanEffectGroup {
  readonly groupIndex: number;
  readonly effect: WorkbookPlanOperationEffect;
  readonly operations: readonly WorkbookPlanEffectOperation[];
}

export interface WorkbookPlanFirstGroupContext {
  /** Complete immutable compiler-owned manifest for every effect group. */
  readonly preconditionManifest?: WorkbookPlanPreconditionManifest;
  readonly executeAfterFresh: (
    adapterObservation: unknown,
    mutation: (group: WorkbookPlanEffectGroup) => Promise<WorkbookPlanGroupExecutionResult>,
  ) => Promise<WorkbookPlanGroupExecutionResult>;
}

export type WorkbookPlanGroupExecutionResult =
  | { readonly status: "committed" }
  | { readonly status: "stale_precondition" }
  | { readonly status: "sync_failed"; readonly failureCode: "sync_failed" };

export interface WorkbookPlanExecutionAdapter {
  readonly executeFirstGroup: (
    context: WorkbookPlanFirstGroupContext,
    signal?: AbortSignal,
  ) => Promise<WorkbookPlanGroupExecutionResult>;
  readonly executeGroup: (
    group: WorkbookPlanEffectGroup,
    groupIndex: number,
    signal?: AbortSignal,
  ) => Promise<WorkbookPlanGroupExecutionResult>;
  readonly reconcileFailedGroup: (
    group: WorkbookPlanEffectGroup,
  ) => Promise<readonly WorkbookPlanOperationState[]>;
  /** Trusted, non-model-visible compiler proofs. Native adapters must enforce these before success. */
  readonly verifyMandatoryPostconditions?: (
    proofs: readonly CompiledWorkbookMandatoryPostcondition[],
    operations: readonly WorkbookPlanEffectOperation[],
    signal?: AbortSignal,
  ) => Promise<boolean>;
  /** Releases adapter-owned proxies/content after this execution reaches a terminal boundary. */
  readonly releaseExecutionState?: (operations: readonly WorkbookPlanEffectOperation[]) => void;
}

export interface IdempotencyStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
  removeItem(key: string): void;
}

export class MemoryIdempotencyStorage implements IdempotencyStorage {
  private readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }

  removeItem(key: string): void {
    this.values.delete(key);
  }
}

interface PersistedRecord {
  readonly boundary: string;
  readonly transactionId: string;
  readonly digest: string;
  readonly owner: string;
  readonly fence: string;
  readonly state: "in_flight" | "terminal";
  readonly createdAt: number;
  readonly report?: WorkbookPlanExecutionReport;
}

interface StoreOptions {
  readonly clock?: () => number;
  readonly ttlMs?: number;
  readonly maxEntries?: number;
  readonly maxBytes?: number;
  readonly coordinator?: IdempotencyCoordinator;
  readonly acceptanceHorizonMs?: number;
}

export interface IdempotencyCoordinator {
  runExclusive<T>(name: string, task: () => Promise<T>): Promise<T>;
}

export type IdempotencyClaim =
  | { readonly kind: "claimed"; readonly owner: string; readonly fence: string }
  | { readonly kind: "terminal"; readonly report: WorkbookPlanExecutionReport }
  | { readonly kind: "conflict" }
  | { readonly kind: "in_flight" }
  | { readonly kind: "restart_rejected" }
  | { readonly kind: "persistence_unavailable" }
  | { readonly kind: "persistence_unavailable_after_mutation"; readonly report: WorkbookPlanExecutionReport };

const STORAGE_KEY = "hermes.excel.workbookPlan.idempotency.v1";
const DEFAULT_TTL_MS = 24 * 60 * 60 * 1000;
const DEFAULT_MAX_ENTRIES = 128;
/**
 * Acceptance window (D5): terminal records younger than this are never capacity-evicted, and a foreign
 * in-flight record older than this is marked as an orphan needing reconciliation.
 */
export const ACCEPTANCE_HORIZON_MS = 10 * 60 * 1000;

interface ReportBinding {
  readonly transactionId: string;
  readonly operationCount: number;
  readonly operations: readonly Pick<WorkbookPlanOperationReport, "operationId" | "operationIndex" | "effect" | "groupIndex">[];
}

export class DeterministicIdempotencyCoordinator implements IdempotencyCoordinator {
  private readonly queues = new Map<string, Promise<void>>();

  async runExclusive<T>(name: string, task: () => Promise<T>): Promise<T> {
    const previous = this.queues.get(name) ?? Promise.resolve();
    let release!: () => void;
    const current = new Promise<void>((resolve) => { release = resolve; });
    this.queues.set(name, current);
    await previous;
    try {
      return await task();
    } finally {
      release();
      if (this.queues.get(name) === current) this.queues.delete(name);
    }
  }
}

class WebLocksIdempotencyCoordinator implements IdempotencyCoordinator {
  async runExclusive<T>(name: string, task: () => Promise<T>): Promise<T> {
    const locks = typeof navigator === "undefined" ? undefined : navigator.locks;
    if (locks === undefined) throw new Error("Web Locks unavailable");
    const result = await locks.request(name, { mode: "exclusive", ifAvailable: true }, async (lock) => {
      if (lock === null) throw new Error("idempotency lock unavailable");
      return task();
    });
    if (result === undefined) throw new Error("idempotency lock unavailable");
    return result;
  }
}

const memoryCoordinators = new WeakMap<object, DeterministicIdempotencyCoordinator>();

function defaultCoordinator(storage: IdempotencyStorage): IdempotencyCoordinator {
  try {
    if (typeof localStorage !== "undefined" && storage === localStorage) return new WebLocksIdempotencyCoordinator();
  } catch {
    return new WebLocksIdempotencyCoordinator();
  }
  const objectStorage = storage as object;
  const existing = memoryCoordinators.get(objectStorage);
  if (existing !== undefined) return existing;
  const coordinator = new DeterministicIdempotencyCoordinator();
  memoryCoordinators.set(objectStorage, coordinator);
  return coordinator;
}
const DEFAULT_MAX_BYTES = 256 * 1024;

export class BrowserIdempotencyStore {
  private readonly clock: () => number;
  private readonly ttlMs: number;
  private readonly maxEntries: number;
  private readonly maxBytes: number;
  private readonly horizonMs: number;
  private readonly ownedInFlight = new Map<string, { readonly owner: string; readonly fence: string; readonly transactionId: string }>();
  private readonly failedFinishReports = new Map<string, WorkbookPlanExecutionReport>();
  private readonly runtimeOwner: string;
  private readonly coordinator: IdempotencyCoordinator;

  constructor(
    private readonly storage: IdempotencyStorage = browserStorage(),
    options: StoreOptions = {},
  ) {
    this.clock = options.clock ?? (() => Date.now());
    this.ttlMs = boundedNonNegative(options.ttlMs ?? DEFAULT_TTL_MS, DEFAULT_TTL_MS);
    this.maxEntries = boundedPositiveInteger(options.maxEntries ?? DEFAULT_MAX_ENTRIES, DEFAULT_MAX_ENTRIES);
    this.maxBytes = boundedPositiveInteger(options.maxBytes ?? DEFAULT_MAX_BYTES, DEFAULT_MAX_BYTES);
    this.horizonMs = boundedNonNegative(options.acceptanceHorizonMs ?? ACCEPTANCE_HORIZON_MS, ACCEPTANCE_HORIZON_MS);
    this.runtimeOwner = freshOpaqueUuid();
    this.coordinator = options.coordinator ?? defaultCoordinator(storage);
  }

  async claim(boundary: string, transactionId: string, digest: string, expected?: ReportBinding): Promise<IdempotencyClaim> {
    if (!isBoundedIdentity(boundary, transactionId, digest)) return { kind: "persistence_unavailable" };
    // The trusted recovery boundary is the durable uniqueness key. The
    // transaction UUID is derived from the digest, so keying by it would let
    // a changed plan evade the replay fence by presenting a new UUID.
    const identityKey = this.key(boundary);
    try {
      return await this.coordinator.runExclusive(STORAGE_KEY, async () => {
        const failedFinishReport = this.failedFinishReports.get(identityKey);
        if (failedFinishReport !== undefined) {
          return { kind: "persistence_unavailable_after_mutation", report: cloneReport(failedFinishReport) };
        }
        const now = this.clock();
        if (!Number.isFinite(now) || now < 0) return { kind: "persistence_unavailable" };
        const loaded = this.loadRecords();
        if (!loaded.ok) return { kind: "persistence_unavailable" };
        const records = loaded.records;
        // Orphan marking (D4): a foreign in-flight claim older than the acceptance window is rewritten once
        // into a terminal restart_rejected marker. createdAt becomes the marking time, so the window for
        // capacity eviction and the TTL both run from marking.
        let marked = false;
        const reconciled = records.map((record): PersistedRecord => {
          if (record.state !== "in_flight" || record.owner === this.runtimeOwner || !(now - record.createdAt > this.horizonMs)) {
            return record;
          }
          marked = true;
          return {
            boundary: record.boundary,
            transactionId: record.transactionId,
            digest: record.digest,
            owner: record.owner,
            fence: record.fence,
            state: "terminal",
            createdAt: now,
            report: {
              outcome: "restart_rejected",
              transactionId: record.transactionId,
              operationCount: 0,
              operations: [],
              verification: "not_attempted",
              rollback: "not_attempted",
            },
          };
        });
        const retained = reconciled.filter((record) =>
          record.state === "in_flight" || now - record.createdAt <= this.ttlMs,
        );
        if ((retained.length !== records.length || marked) && !this.writeBoundedRecords(retained, true, now)) {
          return { kind: "persistence_unavailable" };
        }
        const existing = retained.find((record) => record.boundary === boundary);
        if (existing !== undefined) {
          if (isReconciliationMarker(existing)) return { kind: "restart_rejected" };
          if (existing.digest !== digest) return { kind: "conflict" };
          if (existing.transactionId !== transactionId) return { kind: "conflict" };
          if (existing.state === "terminal" && existing.report !== undefined) {
            if (expected !== undefined && !isReportCompatible(existing.report, expected)) return { kind: "persistence_unavailable" };
            return { kind: "terminal", report: cloneReport(existing.report) };
          }
          if (this.ownedInFlight.has(identityKey)) return { kind: "in_flight" };
          return { kind: "restart_rejected" };
        }
        const fence = freshOpaqueUuid();
        const next: PersistedRecord = {
          boundary, transactionId, digest, owner: this.runtimeOwner, fence,
          state: "in_flight", createdAt: now,
        };
        if (!this.writeBoundedRecords([...retained, next], true, now)) return { kind: "persistence_unavailable" };
        this.ownedInFlight.set(identityKey, { owner: this.runtimeOwner, fence, transactionId });
        return { kind: "claimed", owner: this.runtimeOwner, fence };
      });
    } catch {
      return { kind: "persistence_unavailable" };
    }
  }

  async finish(
    boundary: string,
    transactionId: string,
    digest: string,
    report: WorkbookPlanExecutionReport,
    fence: string,
  ): Promise<boolean> {
    const identityKey = this.key(boundary);
    try {
      return await this.coordinator.runExclusive(STORAGE_KEY, async () => {
        const owned = this.ownedInFlight.get(identityKey);
        const now = this.clock();
        if (!Number.isFinite(now) || now < 0) { this.rememberFinishFailure(identityKey, report, owned?.transactionId); return false; }
        let snapshot: unknown; try { snapshot = cloneReport(report); } catch { snapshot = undefined; }
        const loaded = this.loadRecords();
        // A restart_rejected terminal is a reconciliation marker; it is never written through finish.
        if (!loaded.ok || !isReport(snapshot) || snapshot.outcome === "restart_rejected" || snapshot.transactionId !== transactionId) {
          this.rememberFinishFailure(identityKey, report, owned?.transactionId);
          return false;
        }
        const records = loaded.records;
        const index = records.findIndex((record) =>
          record.boundary === boundary && record.transactionId === transactionId && record.digest === digest &&
          record.state === "in_flight" && record.owner === this.runtimeOwner && record.fence === owned?.fence &&
          record.fence === fence);
        if (index < 0 || owned === undefined) {
          this.rememberFinishFailure(identityKey, report, owned?.transactionId);
          return false;
        }
        const terminal: PersistedRecord = {
          boundary, transactionId, digest, owner: this.runtimeOwner, fence: owned.fence,
          state: "terminal", createdAt: now, report: snapshot,
        };
        const next = [...records];
        next[index] = terminal;
        if (!this.writeBoundedRecords(next, false, now, identityKey)) {
          this.rememberFinishFailure(identityKey, report, owned?.transactionId);
          return false;
        }
        this.ownedInFlight.delete(identityKey);
        this.failedFinishReports.delete(identityKey);
        return true;
      });
    } catch {
      this.rememberFinishFailure(identityKey, report, this.ownedInFlight.get(identityKey)?.transactionId);
      return false;
    }
  }

  /** Number of stored reconciliation markers (content-free; UI deferred to R4). Never marks or writes. */
  async reconciliationRequiredCount(): Promise<number | undefined> {
    try {
      return await this.coordinator.runExclusive(STORAGE_KEY, async () => {
        const loaded = this.loadRecords();
        if (!loaded.ok) return undefined;
        return loaded.records.filter(isReconciliationMarker).length;
      });
    } catch {
      return undefined;
    }
  }

  private key(boundary: string): string {
    return `${this.runtimeOwner}\u0000${boundary}`;
  }

  private loadRecords(): { readonly ok: true; readonly records: readonly PersistedRecord[] } | { readonly ok: false } {
    let raw: string | null;
    try {
      raw = this.storage.getItem(STORAGE_KEY);
      if (raw === null) return { ok: true, records: [] };
      if (typeof raw !== "string" || utf8Length(raw) > this.maxBytes) return { ok: false };
      const parsed: unknown = JSON.parse(raw);
      if (!Array.isArray(parsed) || parsed.length > this.maxEntries) return { ok: false };
      if (!parsed.every(isPersistedRecord)) return { ok: false };
      const boundaries = new Set<string>();
      for (const record of parsed) {
        if (boundaries.has(record.boundary)) return { ok: false };
        boundaries.add(record.boundary);
      }
      return { ok: true, records: parsed };
    } catch {
      return { ok: false };
    }
  }

  /**
   * R4j (R4e-2 lane-A F1 / lane-B A3): when this runtime owns the claim, the remembered report carries the
   * transactionId the claim recorded, never the caller's report id. Without an owned claim the base content is kept.
   */
  private rememberFinishFailure(key: string, report: WorkbookPlanExecutionReport, recordTransactionId: string | undefined): void {
    try {
      const failure = persistenceFailureReport(report);
      this.failedFinishReports.set(key, recordTransactionId === undefined ? failure : { ...failure, transactionId: recordTransactionId });
    } catch {
      let transactionId = "rejected";
      if (recordTransactionId !== undefined) transactionId = recordTransactionId;
      else try { if (isUuid(report.transactionId)) transactionId = report.transactionId; } catch { /* fixed fallback */ }
      this.failedFinishReports.set(key, {
        outcome: "persistence_unavailable_after_mutation",
        transactionId,
        operationCount: 0,
        operations: [],
        verification: "not_attempted",
        rollback: "not_attempted",
      });
    }
  }

  private writeRecords(records: readonly PersistedRecord[]): boolean {
    try {
      if (records.length > this.maxEntries) return false;
      const serialized = JSON.stringify(records);
      if (utf8Length(serialized) > this.maxBytes) return false;
      this.storage.setItem(STORAGE_KEY, serialized);
      return true;
    } catch {
      return false;
    }
  }

  private writeBoundedRecords(
    records: readonly PersistedRecord[],
    preserveInFlight: boolean,
    now: number,
    protectedKey?: string,
  ): boolean {
    const retained = [...records];
    while (retained.length > this.maxEntries || utf8Length(JSON.stringify(retained)) > this.maxBytes) {
      // Only terminal records (markers included) older than the acceptance window may be evicted.
      const removableIndex = retained.findIndex((record) =>
        record.state === "terminal" &&
        now - record.createdAt > this.horizonMs &&
        (protectedKey === undefined || this.key(record.boundary) !== protectedKey),
      );
      if (removableIndex < 0) return false;
      retained.splice(removableIndex, 1);
    }
    return this.writeRecords(retained);
  }
}

interface InternalOperation extends WorkbookPlanEffectOperation {
  readonly effect: WorkbookPlanOperationEffect;
  readonly requiredVerificationKinds: readonly VerificationRule["kind"][];
}

interface InternalAcceptedTransaction {
  readonly publicTransaction: AcceptedWorkbookPlanTransaction;
  readonly trustedOperationBoundary: TrustedOperationBoundary;
  readonly boundaryToken: string;
  readonly groups: readonly WorkbookPlanEffectGroup[];
  readonly operations: readonly InternalOperation[];
  readonly operationDigests: readonly string[];
  readonly postconditions: readonly VerificationRule[];
  readonly mandatoryPostconditions: readonly CompiledWorkbookMandatoryPostcondition[];
  preconditionManifest?: WorkbookPlanPreconditionManifest;
  readonly hasPreconditionManifest: boolean;
  readonly preconditionCheck?: () => Promise<WorkbookPlanPreconditionResult>;
  readonly inverse?: WorkbookPlanInverse;
}

const acceptedTransactions = new WeakMap<object, InternalAcceptedTransaction>();

export async function prepareWorkbookPlanExecution(
  input: WorkbookPlanPreparationInput,
): Promise<WorkbookPlanPreparationResult> {
  const boundaryMetadata = isTrustedOperationBoundary(input.trustedOperationBoundary);
  if (boundaryMetadata === undefined) {
    return { accepted: false, reason: "invalid_boundary" };
  }

  let canonical: CanonicalizeWorkbookPlanSuccess;
  try {
    const result = await canonicalizeWorkbookPlan(input.plan);
    if (!result.ok) return { accepted: false, reason: "canonicalization_failed" };
    canonical = result;
  } catch {
    return { accepted: false, reason: "canonicalization_failed" };
  }

  const snapshot = createWorkbookPlanPolicyInputSnapshot(input.policyInput);
  const policy = decideWorkbookPlanPolicy(snapshot);
  if (policy.decision !== "approved") {
    return { accepted: false, reason: "policy_rejected", policyReason: policy.reasonCode };
  }

  let compiled: ReturnType<typeof compileWorkbookPlan>;
  try {
    compiled = compileWorkbookPlan(canonical.value, input.compileContext);
  } catch {
    return { accepted: false, reason: "compile_failed" };
  }
  if (!compiled.ok) return { accepted: false, reason: "compile_failed" };
  if (compiled.transaction.mandatoryProofsComplete !== true) {
    return { accepted: false, reason: "compile_failed" };
  }

  let transactionId: string;
  let operationIds: readonly string[];
  try {
    if (input.uuidSource !== undefined) {
      transactionId = nextUuid(input.uuidSource);
      operationIds = compiled.transaction.primitives.map(() => nextUuid(input.uuidSource!));
    } else {
      transactionId = await deriveDeterministicUuid(
        "hermes.excel.workbook-plan.transaction.v1",
        boundaryMetadata.token,
        canonical.semanticDigest,
      );
      operationIds = await Promise.all(compiled.transaction.primitives.map((primitive) => deriveDeterministicUuid(
        "hermes.excel.workbook-plan.operation.v1",
        boundaryMetadata.token,
        canonical.semanticDigest,
        primitive.operationIndex,
      )));
    }
  } catch {
    return { accepted: false, reason: "invalid_uuid_source" };
  }
  const operations: InternalOperation[] = compiled.transaction.primitives.map((primitive, index) => ({
    operationId: operationIds[index],
    operationIndex: primitive.operationIndex,
    kind: primitive.kind,
    primitiveName: primitive.primitiveName,
    args: deepFreeze(cloneContent(primitive.args)) as Readonly<Record<string, unknown>>,
    effect: primitive.effect,
    requiredVerificationKinds: primitive.requiredVerificationKinds,
  }));
  const groups = groupOperations(operations);
  if (
    canonical.operationDigests.length !== operations.length ||
    canonical.operationDigests.some((digest) => !/^[0-9a-f]{64}$/.test(digest))
  ) {
    return { accepted: false, reason: "canonicalization_failed" };
  }
  const mandatoryPostconditions = (compiled.transaction.mandatoryPostconditions ?? []).map((entry) => ({
    ...entry,
    ...(entry.proof === undefined ? {} : {
      proof: {
        ...entry.proof,
        argsDigest: canonical.operationDigests[entry.operationIndex],
      },
    }),
  }));
  const publicTransaction: AcceptedWorkbookPlanTransaction = Object.freeze({
    transactionId,
    semanticDigest: canonical.semanticDigest,
    operationCount: operations.length,
  });
  const internal: InternalAcceptedTransaction = {
    publicTransaction,
    trustedOperationBoundary: input.trustedOperationBoundary,
    boundaryToken: boundaryMetadata.token,
    groups,
    operations,
    operationDigests: deepFreeze([...canonical.operationDigests]) as readonly string[],
    postconditions: deepFreeze(cloneContent(compiled.transaction.postconditions)) as readonly VerificationRule[],
    mandatoryPostconditions: deepFreeze(cloneContent(mandatoryPostconditions)) as readonly CompiledWorkbookMandatoryPostcondition[],
    preconditionManifest: input.compileContext.preconditionManifest === undefined
      ? undefined
      : deepFreeze(cloneContent(input.compileContext.preconditionManifest)) as WorkbookPlanPreconditionManifest,
    hasPreconditionManifest: input.compileContext.preconditionManifest !== undefined,
    preconditionCheck: input.preconditionCheck,
    inverse: input.inverse === undefined ? undefined : Object.freeze({
      completeCapturedPreState: input.inverse.completeCapturedPreState === true,
      independentlyVerifiedSafe: input.inverse.independentlyVerifiedSafe === true,
      rollback: input.inverse.rollback,
    }),
  };
  acceptedTransactions.set(publicTransaction, internal);
  return { accepted: true, transaction: publicTransaction };
}

export async function executeAcceptedWorkbookPlan(
  transaction: AcceptedWorkbookPlanTransaction,
  trustedOperationBoundary: TrustedOperationBoundary,
  adapter: WorkbookPlanExecutionAdapter,
  store: BrowserIdempotencyStore,
  options: WorkbookPlanExecutionOptions | AbortSignal = {},
): Promise<WorkbookPlanExecutionReport> {
  const internal = acceptedTransactions.get(transaction);
  if (
    internal === undefined ||
    isTrustedOperationBoundary(trustedOperationBoundary) === undefined ||
    isTrustedOperationBoundary(trustedOperationBoundary)?.token !== internal.boundaryToken
  ) {
    return rejectedReport(transaction);
  }

  // A4: capture options before the claim, so a malformed callback never consumes the boundary.
  const capturedOptions = captureExecutionOptions(options);
  if (capturedOptions.malformedBeforeFirstMutation || capturedOptions.malformedDeadlineAt) return baseReport(internal, "cancelled");

  const claim = await store.claim(
    internal.boundaryToken,
    internal.publicTransaction.transactionId,
    internal.publicTransaction.semanticDigest,
    reportBinding(internal),
  );
  if (claim.kind === "terminal") {
    replayedReports.add(claim.report);
    return claim.report;
  }
  if (claim.kind === "conflict") return baseReport(internal, "conflict");
  if (claim.kind === "restart_rejected") return baseReport(internal, "restart_rejected");
  if (claim.kind === "in_flight") return baseReport(internal, "restart_rejected");
  if (claim.kind === "persistence_unavailable") return baseReport(internal, "persistence_unavailable");
  if (claim.kind === "persistence_unavailable_after_mutation") return claim.report;

  const report = await executeClaimedSafely(internal, adapter, capturedOptions);
  const durable = await store.finish(
    internal.boundaryToken,
    internal.publicTransaction.transactionId,
    internal.publicTransaction.semanticDigest,
    report,
    claim.fence,
  );
  return durable ? report : persistenceFailureReport(report);
}

/** Reports that executeAcceptedWorkbookPlan returned from a terminal claim (an earlier attempt's recorded report). */
const replayedReports = new WeakSet<object>();

/** R4e: true only for a report returned as a terminal-claim replay; the report's bytes are unchanged. */
export function isReplayedWorkbookPlanReport(value: unknown): boolean {
  return typeof value === "object" && value !== null && replayedReports.has(value);
}

export function getAcceptedWorkbookPlanPostconditionsForVerification(
  transaction: AcceptedWorkbookPlanTransaction,
): readonly VerificationRule[] | undefined {
  return acceptedTransactions.get(transaction)?.postconditions;
}

export interface WorkbookPlanExecutionOptions {
  readonly signal?: AbortSignal;
  readonly beforeFirstMutation?: () => Promise<void>;
  readonly deadlineAt?: number;
}

interface CapturedWorkbookPlanExecutionOptions {
  readonly signal?: AbortSignal;
  readonly beforeFirstMutation?: () => Promise<void>;
  readonly malformedBeforeFirstMutation: boolean;
  readonly deadlineAt?: number;
  readonly malformedDeadlineAt: boolean;
}

function captureExecutionOptions(
  options: WorkbookPlanExecutionOptions | AbortSignal,
): CapturedWorkbookPlanExecutionOptions {
  try {
    if (typeof AbortSignal !== "undefined" && options instanceof AbortSignal) {
      return { signal: options, malformedBeforeFirstMutation: false, malformedDeadlineAt: false };
    }
  } catch {
    return { malformedBeforeFirstMutation: false, malformedDeadlineAt: false };
  }

  let signal: AbortSignal | undefined;
  try {
    signal = (options as WorkbookPlanExecutionOptions).signal;
  } catch {
    signal = undefined;
  }

  // R4e GA3: read exactly once; absent stays absent, a finite number is captured, anything else is malformed.
  let deadline: { readonly deadlineAt?: number; readonly malformedDeadlineAt: boolean };
  try {
    const deadlineAt: unknown = (options as WorkbookPlanExecutionOptions).deadlineAt;
    if (deadlineAt === undefined) deadline = { malformedDeadlineAt: false };
    else if (typeof deadlineAt === "number" && Number.isFinite(deadlineAt)) deadline = { deadlineAt, malformedDeadlineAt: false };
    else deadline = { malformedDeadlineAt: true };
  } catch {
    deadline = { malformedDeadlineAt: true };
  }

  let beforeFirstMutation: unknown;
  try {
    beforeFirstMutation = (options as WorkbookPlanExecutionOptions).beforeFirstMutation;
  } catch {
    return { signal, malformedBeforeFirstMutation: true, ...deadline };
  }
  if (beforeFirstMutation === undefined) return { signal, malformedBeforeFirstMutation: false, ...deadline };
  if (typeof beforeFirstMutation !== "function") return { signal, malformedBeforeFirstMutation: true, ...deadline };
  return {
    signal,
    beforeFirstMutation: beforeFirstMutation as () => Promise<void>,
    malformedBeforeFirstMutation: false,
    ...deadline,
  };
}

async function executeClaimedSafely(
  internal: InternalAcceptedTransaction,
  adapter: WorkbookPlanExecutionAdapter,
  options: CapturedWorkbookPlanExecutionOptions,
): Promise<WorkbookPlanExecutionReport> {
  try {
    if (options.malformedBeforeFirstMutation || options.malformedDeadlineAt) return baseReport(internal, "cancelled");
    return await executeClaimed(internal, adapter, options.signal, options.beforeFirstMutation, options.deadlineAt);
  } catch {
    return conservativeFailureReport(internal);
  } finally {
    internal.preconditionManifest = undefined;
    try {
      adapter.releaseExecutionState?.(internal.groups.flatMap((group) => group.operations));
    } catch { /* cleanup is best effort; the execution result remains fail-closed */ }
  }
}

async function executeClaimed(
  internal: InternalAcceptedTransaction,
  adapter: WorkbookPlanExecutionAdapter,
  signal?: AbortSignal,
  beforeFirstMutation?: () => Promise<void>,
  deadlineAt?: number,
): Promise<WorkbookPlanExecutionReport> {
  const mutable = internal.operations.map((operation, operationIndex) => ({
    operation,
    groupIndex: groupIndexFor(internal.groups, operationIndex),
    state: "not_committed" as WorkbookPlanOperationState,
    status: "not_started" as WorkbookPlanOperationStatus,
  }));
  let verification: WorkbookPlanVerificationStatus = "not_attempted";
  let rollback: WorkbookPlanRollbackStatus = "not_attempted";
  let firstFailedGroup: number | undefined;
  let firstFailedOperationIndex: number | undefined;
  let outcome: WorkbookPlanExecutionOutcome = "committed";
  let stopAfterAbort = false;
  let firstGroupHandedOff = false;
  let firstGroupAuthorizationRejected = false;

  for (let index = 0; index < internal.groups.length; index += 1) {
    const group = internal.groups[index];
    if (safeAborted(signal)) { stopAfterAbort = true; break; }
    let result: WorkbookPlanGroupExecutionResult;
    try {
      if (index === 0) {
        let gateUsed = false;
        let gateRevoked = false;
        let gateMisused = false;
        let mutationRejected = false;
        let gatePromise: Promise<WorkbookPlanGroupExecutionResult> | undefined;
        const context: WorkbookPlanFirstGroupContext = {
          get preconditionManifest() { return internal.preconditionManifest; },
          executeAfterFresh: (observation, mutation) => {
            if (gateUsed) {
              gateMisused = true;
              return Promise.resolve({ status: "sync_failed", failureCode: "sync_failed" } as const);
            }
            if (gateRevoked || safeAborted(signal)) {
              return Promise.resolve({ status: "sync_failed", failureCode: "sync_failed" } as const);
            }
            gateUsed = true;
            gatePromise = (async () => {
              if (observation !== "fresh" && observation !== "stale") return { status: "sync_failed", failureCode: "sync_failed" };
              if (observation === "stale") return { status: "stale_precondition" };
              let trusted: unknown = "fresh";
              if (internal.preconditionCheck !== undefined) {
                try { trusted = await internal.preconditionCheck(); } catch { return { status: "stale_precondition" }; }
              }
              if (gateRevoked) return { status: "sync_failed", failureCode: "sync_failed" };
              if (trusted === "stale") return { status: "stale_precondition" };
              if (trusted !== "fresh" || safeAborted(signal)) return { status: "sync_failed", failureCode: "sync_failed" };
              if (beforeFirstMutation !== undefined) {
                try {
                  await beforeFirstMutation();
                } catch {
                  firstGroupAuthorizationRejected = true;
                  return { status: "sync_failed", failureCode: "sync_failed" };
                }
                if (gateRevoked || safeAborted(signal)) {
                  firstGroupAuthorizationRejected = true;
                  return { status: "sync_failed", failureCode: "sync_failed" };
                }
              }
              if (deadlineAt !== undefined && !(Date.now() < deadlineAt)) {
                firstGroupAuthorizationRejected = true;
                return { status: "sync_failed", failureCode: "sync_failed" };
              }
              firstGroupHandedOff = true;
              try { return normalizeGroupExecutionResult(await mutation(group)); } catch {
                mutationRejected = true;
                return { status: "sync_failed", failureCode: "sync_failed" };
              }
            })();
            return gatePromise.then((result) => {
              if (mutationRejected) throw new Error("First-group mutation rejected");
              return result;
            });
          },
        };
        try {
          await adapter.executeFirstGroup(context, signal);
        } catch { /* the exact gated result below is the only first-group authority */ }
        gateRevoked = true;
        if (!gateUsed || gatePromise === undefined) {
          result = { status: "sync_failed", failureCode: "sync_failed" };
        } else {
          try {
            result = normalizeGroupExecutionResult(await gatePromise);
          } catch {
            result = { status: "sync_failed", failureCode: "sync_failed" };
          }
          if (gateMisused) result = { status: "sync_failed", failureCode: "sync_failed" };
        }
        internal.preconditionManifest = undefined;
      } else {
        if (safeAborted(signal)) {
          stopAfterAbort = true;
          break;
        }
        result = normalizeGroupExecutionResult(await adapter.executeGroup(group, index, signal));
      }
    } catch {
      result = { status: "sync_failed", failureCode: "sync_failed" };
    }
    if (index === 0 && safeAborted(signal) && !firstGroupHandedOff) {
      stopAfterAbort = true;
      break;
    }
    if (index === 0 && firstGroupAuthorizationRejected) {
      stopAfterAbort = true;
      break;
    }
    if (index > 0 && result.status === "stale_precondition") result = { status: "sync_failed", failureCode: "sync_failed" };
    if (result.status === "committed") {
      setGroupState(mutable, group, "committed", "applied");
      if (safeAborted(signal)) { stopAfterAbort = true; break; }
      continue;
    }
    if (result.status === "stale_precondition" && index === 0) {
      firstFailedGroup = index;
      firstFailedOperationIndex = group.operations[0]?.operationIndex;
      outcome = "stale";
      break;
    }
    firstFailedGroup = index;
    firstFailedOperationIndex = group.operations[0]?.operationIndex;
    let reconciled: readonly WorkbookPlanOperationState[] = group.operations.map(() => "unknown_within_batch" as WorkbookPlanOperationState);
    try {
      const normalized = normalizeReconciliation(await adapter.reconcileFailedGroup(group), group.operations.length);
      if (normalized === undefined) verification = "reconciliation_failed";
      else reconciled = normalized;
    } catch { verification = "reconciliation_failed"; }
    if (verification === "not_attempted") verification = "reconciled";
    for (let operationIndex = 0; operationIndex < group.operations.length; operationIndex += 1) {
      const state = reconciled[operationIndex];
      const item = mutable.find((candidate) => candidate.operation.operationIndex === group.operations[operationIndex].operationIndex);
      if (item !== undefined) { item.state = state; item.status = statusForState(state); }
    }
    // Cancellation after an Office handoff is not evidence that an inverse is
    // safe. Reconciliation is recovery for the unresolved batch; do not turn
    // that recovery path into a rollback claim.
    if (!safeAborted(signal)) rollback = await attemptRollback(internal.inverse);
    break;
  }
  if (firstFailedGroup === undefined) {
    outcome = stopAfterAbort ? (mutable.some((item) => item.state !== "not_committed") ? "partial" : "cancelled") : "committed";
  } else if (outcome === "stale") {
    for (const item of mutable) { item.state = "not_committed"; item.status = "not_started"; }
  } else outcome = "partial";

  const mandatoryProofs = !internal.hasPreconditionManifest
    ? { ok: true as const, proofs: [] as readonly CompiledWorkbookMandatoryPostcondition[] }
    : validateMandatoryProofs(internal);
  if (outcome === "committed" && internal.hasPreconditionManifest) {
    if (!mandatoryProofs.ok || safeAborted(signal) || adapter.verifyMandatoryPostconditions === undefined) {
      outcome = "partial";
      verification = "reconciliation_failed";
      markProofOperationsUnknown(mutable, internal.mandatoryPostconditions);
    } else {
      let proven = false;
      try {
        proven = await adapter.verifyMandatoryPostconditions(mandatoryProofs.proofs, internal.operations, signal);
      } catch { proven = false; }
      if (!proven) {
        outcome = "partial";
        verification = "reconciliation_failed";
        markProofOperationsUnknown(mutable, mandatoryProofs.proofs);
      }
    }
  }
  if (rollback === "succeeded") {
    for (const item of mutable) if (item.state === "committed" || item.state === "unknown_within_batch") item.status = "rolled_back";
  }
  return {
    outcome,
    transactionId: internal.publicTransaction.transactionId,
    operationCount: internal.operations.length,
    operations: mutable.map(({ operation, groupIndex, state, status }) => ({ operationId: operation.operationId, operationIndex: operation.operationIndex, effect: operation.effect, groupIndex, state, status })),
    ...(firstFailedGroup === undefined ? {} : { firstFailedGroup }),
    ...(firstFailedOperationIndex === undefined ? {} : { firstFailedOperationIndex }),
    verification,
    rollback,
  };
}

function conservativeFailureReport(internal: InternalAcceptedTransaction): WorkbookPlanExecutionReport {
  const report = baseReport(internal, "partial");
  const first = internal.groups[0];
  if (first === undefined) return { ...report, verification: "reconciliation_failed" };
  return {
    ...report,
    firstFailedGroup: 0,
    firstFailedOperationIndex: first.operations[0]?.operationIndex,
    operations: report.operations.map((operation) => first.operations.some((item) => item.operationIndex === operation.operationIndex)
      ? { ...operation, state: "unknown_within_batch", status: "unknown_within_batch" } : operation),
    verification: "reconciliation_failed",
  };
}

async function attemptRollback(inverse: WorkbookPlanInverse | undefined): Promise<WorkbookPlanRollbackStatus> {
  if (inverse === undefined || !inverse.completeCapturedPreState || !inverse.independentlyVerifiedSafe) {
    return "unavailable";
  }
  try {
    return (await inverse.rollback()) ? "succeeded" : "failed";
  } catch {
    return "failed";
  }
}

function groupOperations(operations: readonly InternalOperation[]): readonly WorkbookPlanEffectGroup[] {
  const groups: WorkbookPlanEffectGroup[] = [];
  for (const operation of operations) {
    const previous = groups[groups.length - 1];
    const item: WorkbookPlanEffectOperation = operation;
    if (previous !== undefined && previous.effect === operation.effect) {
      groups[groups.length - 1] = Object.freeze({
        ...previous,
        operations: Object.freeze([...previous.operations, item]),
      });
    } else {
      groups.push(Object.freeze({
        groupIndex: groups.length,
        effect: operation.effect,
        operations: Object.freeze([item]),
      }));
    }
  }
  return Object.freeze(groups);
}

function groupIndexFor(groups: readonly WorkbookPlanEffectGroup[], operationIndex: number): number {
  for (const group of groups) {
    if (group.operations.some((operation) => operation.operationIndex === operationIndex)) return group.groupIndex;
  }
  return 0;
}

function setGroupState(
  mutable: Array<{ readonly operation: InternalOperation; groupIndex: number; state: WorkbookPlanOperationState; status: WorkbookPlanOperationStatus }>,
  group: WorkbookPlanEffectGroup,
  state: WorkbookPlanOperationState,
  status: WorkbookPlanOperationStatus,
): void {
  for (const operation of group.operations) {
    const item = mutable.find((candidate) => candidate.operation.operationIndex === operation.operationIndex);
    if (item !== undefined) {
      item.state = state;
      item.status = status;
    }
  }
}

function markProofOperationsUnknown(
  mutable: Array<{ readonly operation: InternalOperation; groupIndex: number; state: WorkbookPlanOperationState; status: WorkbookPlanOperationStatus }>,
  proofs: readonly CompiledWorkbookMandatoryPostcondition[],
): void {
  const indexes = new Set(proofs.map((proof) => proof.operationIndex));
  for (const item of mutable) {
    if (indexes.has(item.operation.operationIndex)) {
      item.state = "unknown_within_batch";
      item.status = "unknown_within_batch";
    }
  }
}

function validateMandatoryProofs(
  internal: InternalAcceptedTransaction,
): { readonly ok: true; readonly proofs: readonly CompiledWorkbookMandatoryPostcondition[] } | { readonly ok: false } {
  if (!validateTrustedOperationProofBindings(internal.mandatoryPostconditions, internal.operations, internal.operationDigests)) return { ok: false };
  const byIndex = new Map<number, CompiledWorkbookMandatoryPostcondition>();
  for (const entry of internal.mandatoryPostconditions) {
    const operation = internal.operations.find((candidate) => candidate.operationIndex === entry.operationIndex);
    if (
      operation === undefined ||
      entry.proof === undefined ||
      entry.proof.operationIndex !== entry.operationIndex ||
      entry.proof.operationKind !== operation.kind ||
      entry.proof.primitiveName !== operation.primitiveName ||
      entry.proof.effect !== operation.effect ||
      entry.proof.argsDigest !== internal.operationDigests[entry.operationIndex] ||
      operation.operationIndex !== entry.operationIndex ||
      !operationKindProofMatches(entry, operation)
    ) return { ok: false };
    if (byIndex.has(entry.operationIndex)) return { ok: false };
    byIndex.set(entry.operationIndex, entry);
  }
  const proofs: CompiledWorkbookMandatoryPostcondition[] = [];
  for (let index = 0; index < internal.operations.length; index += 1) {
    const entry = byIndex.get(internal.operations[index]!.operationIndex);
    if (entry === undefined) return { ok: false };
    proofs.push(entry);
  }
  return { ok: true, proofs };
}

/**
 * Validate the immutable trusted proof-to-primitive binding independently of
 * the host readback. This is exported only for direct executor contract tests;
 * native execution still calls it through validateMandatoryProofs().
 */
export function validateTrustedOperationProofBindings(
  proofs: readonly CompiledWorkbookMandatoryPostcondition[],
  operations: readonly (WorkbookPlanEffectOperation & { readonly effect: string })[],
  operationDigests: readonly string[],
): boolean {
  if (proofs.length !== operations.length || operationDigests.length !== operations.length ||
    operationDigests.some((digest) => !/^[0-9a-f]{64}$/.test(digest))) return false;
  const byIndex = new Map<number, CompiledWorkbookEffectOperationForBinding>();
  for (const operation of operations) {
    if (byIndex.has(operation.operationIndex)) return false;
    byIndex.set(operation.operationIndex, {
      operationIndex: operation.operationIndex,
      kind: operation.kind,
      primitiveName: operation.primitiveName,
      effect: operation.effect,
      args: operation.args,
    });
  }
  const proofIndexes = new Set<number>();
  for (const entry of proofs) {
    const operation = byIndex.get(entry.operationIndex);
    const proof = entry.proof;
    if (
      operation === undefined ||
      proof === undefined ||
      proofIndexes.has(entry.operationIndex) ||
      proof.operationIndex !== entry.operationIndex ||
      proof.operationKind !== operation.kind ||
      proof.primitiveName !== operation.primitiveName ||
      proof.effect !== operation.effect ||
      proof.argsDigest !== operationDigests[entry.operationIndex]
    ) return false;
    proofIndexes.add(entry.operationIndex);
  }
  return operations.every((operation) => proofIndexes.has(operation.operationIndex));
}

type CompiledWorkbookEffectOperationForBinding = Pick<WorkbookPlanEffectOperation, "operationIndex" | "kind" | "primitiveName" | "args"> & { readonly effect: string };

function operationKindProofMatches(
  entry: CompiledWorkbookMandatoryPostcondition,
  operation: InternalOperation,
): boolean {
  return operation.kind === entry.proof?.operationKind &&
    operation.requiredVerificationKinds.includes(entry.kind);
}

function statusForState(state: WorkbookPlanOperationState): WorkbookPlanOperationStatus {
  switch (state) {
    case "committed": return "applied";
    case "unknown_within_batch": return "unknown_within_batch";
    case "not_committed": return "failed";
  }
}

function baseReport(
  internal: InternalAcceptedTransaction,
  outcome: WorkbookPlanExecutionOutcome,
): WorkbookPlanExecutionReport {
  return {
    outcome,
    transactionId: internal.publicTransaction.transactionId,
    operationCount: internal.operations.length,
    operations: internal.operations.map((operation) => ({
      operationId: operation.operationId,
      operationIndex: operation.operationIndex,
      effect: operation.effect,
      groupIndex: groupIndexFor(internal.groups, operation.operationIndex),
      state: "not_committed",
      status: "not_started",
    })),
    verification: "not_attempted",
    rollback: "not_attempted",
  };
}

function rejectedReport(transaction: AcceptedWorkbookPlanTransaction): WorkbookPlanExecutionReport {
  const safeTransactionId = isUuid(transaction?.transactionId) ? transaction.transactionId : "rejected";
  const count = Number.isSafeInteger(transaction?.operationCount) && transaction.operationCount >= 0
    ? Math.min(transaction.operationCount, 10000)
    : 0;
  return {
    outcome: "rejected",
    transactionId: safeTransactionId,
    operationCount: count,
    operations: [],
    verification: "not_attempted",
    rollback: "not_attempted",
  };
}

function cloneReport(report: WorkbookPlanExecutionReport): WorkbookPlanExecutionReport {
  return JSON.parse(JSON.stringify(report)) as WorkbookPlanExecutionReport;
}

let fallbackUuidCounter = 0;

function freshOpaqueUuid(): string {
  try {
    const candidate = crypto.randomUUID();
    if (isUuid(candidate)) return candidate;
  } catch { /* fail closed to a bounded process-local token */ }
  fallbackUuidCounter += 1;
  return `00000000-0000-4000-8000-${fallbackUuidCounter.toString(16).padStart(12, "0")}`;
}

async function deriveDeterministicUuid(
  domain: string,
  boundaryToken: string,
  semanticDigest: string,
  operationIndex?: number,
): Promise<string> {
  const material = [domain, boundaryToken, semanticDigest, operationIndex === undefined ? "transaction" : String(operationIndex)].join("\u0000");
  const digest = await globalThis.crypto.subtle.digest("SHA-256", new TextEncoder().encode(material));
  const bytes = new Uint8Array(digest).slice(0, 16);
  // Deterministic SHA-256-derived bytes with the UUID version/variant shape.
  // This is not RFC 4122 UUIDv5, which is defined over SHA-1 material.
  bytes[6] = (bytes[6] & 0x0f) | 0x50;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
  const value = `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  if (!isUuid(value)) throw new Error("Deterministic UUID derivation failed");
  return value;
}

function isTrustedOperationBoundary(value: unknown): TrustedBoundaryMetadata | undefined {
  if (value === null || typeof value !== "object") return undefined;
  return trustedBoundaries.get(value);
}

function reportBinding(internal: InternalAcceptedTransaction): ReportBinding {
  return {
    transactionId: internal.publicTransaction.transactionId,
    operationCount: internal.operations.length,
    operations: internal.operations.map((operation) => ({
      operationId: operation.operationId,
      operationIndex: operation.operationIndex,
      effect: operation.effect,
      groupIndex: groupIndexFor(internal.groups, operation.operationIndex),
    })),
  };
}

function isReportCompatible(report: WorkbookPlanExecutionReport, expected: ReportBinding): boolean {
  try {
    return report.transactionId === expected.transactionId &&
      report.operationCount === expected.operationCount &&
      report.operations.length === expected.operations.length &&
      report.operations.every((operation, index) => {
        const item = expected.operations[index];
        return item !== undefined && operation.operationId === item.operationId &&
          operation.operationIndex === item.operationIndex && operation.effect === item.effect &&
          operation.groupIndex === item.groupIndex && validStateStatus(operation.state, operation.status);
      });
  } catch { return false; }
}

function validStateStatus(state: WorkbookPlanOperationState, status: WorkbookPlanOperationStatus): boolean {
  return (state === "committed" && ["applied", "verified", "rolled_back"].includes(status)) ||
    (state === "not_committed" && ["not_started", "failed"].includes(status)) ||
    (state === "unknown_within_batch" && ["unknown_within_batch", "rolled_back"].includes(status));
}

function validReportInvariants(report: WorkbookPlanExecutionReport): boolean {
  if (!report.operations.every((operation) => validStateStatus(operation.state, operation.status))) return false;
  if (report.outcome === "committed" && report.operations.some((operation) => operation.state !== "committed")) return false;
  if ((report.outcome === "stale" || report.outcome === "cancelled") &&
    report.operations.some((operation) => operation.state !== "not_committed")) return false;
  return true;
}

function safeAborted(signal: AbortSignal | undefined): boolean {
  if (signal === undefined) return false;
  try { return signal.aborted === true; } catch { return true; }
}

function normalizeGroupExecutionResult(value: unknown): WorkbookPlanGroupExecutionResult {
  try {
    if (value === null || typeof value !== "object" || Array.isArray(value)) return { status: "sync_failed", failureCode: "sync_failed" };
    const record = value as Record<string, unknown>;
    const keys = Object.keys(record);
    if (record.status === "committed" && keys.length === 1 && keys[0] === "status") return { status: "committed" };
    if (record.status === "stale_precondition" && keys.length === 1 && keys[0] === "status") return { status: "stale_precondition" };
    if (record.status === "sync_failed" && record.failureCode === "sync_failed" && keys.length === 2 && keys.includes("status") && keys.includes("failureCode")) {
      return { status: "sync_failed", failureCode: "sync_failed" };
    }
  } catch { /* normalize to fixed failure */ }
  return { status: "sync_failed", failureCode: "sync_failed" };
}

function normalizeReconciliation(value: unknown, expectedLength: number): readonly WorkbookPlanOperationState[] | undefined {
  try {
    if (!Array.isArray(value)) return undefined;
    const lengthDescriptor = Object.getOwnPropertyDescriptor(value, "length");
    if (lengthDescriptor === undefined || !("value" in lengthDescriptor) ||
      !Number.isSafeInteger(lengthDescriptor.value) || lengthDescriptor.value !== expectedLength) return undefined;
    const result: WorkbookPlanOperationState[] = [];
    for (let index = 0; index < expectedLength; index += 1) {
      const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
      if (descriptor === undefined || !("value" in descriptor)) return undefined;
      const state = descriptor.value;
      if (!isOperationState(state)) return undefined;
      result.push(state);
    }
    return result;
  } catch { return undefined; }
}

function persistenceFailureReport(report: WorkbookPlanExecutionReport): WorkbookPlanExecutionReport {
  return { ...cloneReport(report), outcome: "persistence_unavailable_after_mutation" };
}

function isReconciliationMarker(record: PersistedRecord): boolean {
  return record.state === "terminal" && record.report?.outcome === "restart_rejected";
}

function isPersistedRecord(value: unknown): value is PersistedRecord {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  const requiredKeys = record.state === "terminal"
    ? ["boundary", "transactionId", "digest", "owner", "fence", "state", "createdAt", "report"]
    : ["boundary", "transactionId", "digest", "owner", "fence", "state", "createdAt"];
  if (!hasOnlyKeys(record, requiredKeys) ||
    !isBoundedIdentity(record.boundary, record.transactionId, record.digest) ||
    !isUuid(record.owner) || !isUuid(record.fence) ||
    !Number.isFinite(record.createdAt) || (record.createdAt as number) < 0 ||
    (record.state !== "in_flight" && record.state !== "terminal")) return false;
  if (record.state === "in_flight") return true;
  return isReport(record.report) && record.report.transactionId === record.transactionId &&
    record.report.operations.every((operation) => validStateStatus(operation.state, operation.status));
}

function isReport(value: unknown): value is WorkbookPlanExecutionReport {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
  const report = value as Record<string, unknown>;
  if (!hasOnlyKeys(report, [
    "outcome", "transactionId", "operationCount", "operations", "verification", "rollback",
  ], ["firstFailedGroup", "firstFailedOperationIndex"])) return false;
  const outcomes: readonly WorkbookPlanExecutionOutcome[] = [
    "committed", "partial", "stale", "cancelled", "replay", "conflict", "restart_rejected", "rejected",
    "persistence_unavailable", "persistence_unavailable_after_mutation",
  ];
  if (!outcomes.includes(report.outcome as WorkbookPlanExecutionOutcome) ||
    !isUuid(report.transactionId) ||
    !Number.isSafeInteger(report.operationCount) || (report.operationCount as number) < 0 ||
    (report.operationCount as number) > 10000 ||
    !Array.isArray(report.operations) || report.operations.length !== report.operationCount ||
    !(["not_attempted", "reconciled", "reconciliation_failed"] as const).includes(report.verification as never) ||
    !(["not_attempted", "unavailable", "succeeded", "failed"] as const).includes(report.rollback as never)) return false;
  if (report.firstFailedGroup !== undefined && (!Number.isSafeInteger(report.firstFailedGroup) || (report.firstFailedGroup as number) < 0)) return false;
  if (report.firstFailedOperationIndex !== undefined && (!Number.isSafeInteger(report.firstFailedOperationIndex) || (report.firstFailedOperationIndex as number) < 0)) return false;
  return report.operations.every(isOperationReport) && validReportInvariants(report as unknown as WorkbookPlanExecutionReport);
}

function isOperationReport(value: unknown): value is WorkbookPlanOperationReport {
  if (value === null || typeof value !== "object" || Array.isArray(value)) return false;
  const operation = value as Record<string, unknown>;
  if (!hasOnlyKeys(operation, ["operationId", "operationIndex", "effect", "groupIndex", "state", "status"])) return false;
  return isUuid(operation.operationId) &&
    Number.isSafeInteger(operation.operationIndex) && (operation.operationIndex as number) >= 0 &&
    (["write", "format", "structure", "filter", "calculate"] as const).includes(operation.effect as never) &&
    Number.isSafeInteger(operation.groupIndex) && (operation.groupIndex as number) >= 0 &&
    (["committed", "not_committed", "unknown_within_batch"] as const).includes(operation.state as never) &&
    (["not_started", "applied", "verified", "failed", "rolled_back", "unknown_within_batch"] as const).includes(operation.status as never);
}

function hasOnlyKeys(
  value: Record<string, unknown>,
  required: readonly string[],
  optional: readonly string[] = [],
): boolean {
  const allowed = new Set([...required, ...optional]);
  return Object.keys(value).every((key) => allowed.has(key)) &&
    required.every((key) => Object.prototype.hasOwnProperty.call(value, key));
}

function isOperationState(value: unknown): value is WorkbookPlanOperationState {
  return value === "committed" || value === "not_committed" || value === "unknown_within_batch";
}

function cloneContent(value: unknown): unknown {
  if (value === null || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map(cloneContent);
  const record = value as Readonly<Record<string, unknown>>;
  const result: Record<string, unknown> = {};
  for (const key of Object.keys(record)) result[key] = cloneContent(record[key]);
  return result;
}

function deepFreeze(value: unknown): unknown {
  if (value === null || typeof value !== "object" || Object.isFrozen(value)) return value;
  if (Array.isArray(value)) {
    for (const item of value) deepFreeze(item);
  } else {
    const record = value as Readonly<Record<string, unknown>>;
    for (const key of Object.keys(record)) deepFreeze(record[key]);
  }
  return Object.freeze(value);
}

function nextUuid(source: WorkbookPlanUuidSource): string {
  const value = source();
  if (!isUuid(value)) throw new Error("UUID source returned an invalid UUID");
  return value;
}

function isUuid(value: unknown): value is string {
  return typeof value === "string" && UUID_PATTERN.test(value);
}

function isBoundedBoundary(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= MAX_BOUNDARY_LENGTH &&
    value.split(".").length === 3 && value.split(".").every(isUuid);
}

function isBoundedIdentity(boundary: unknown, transactionId: unknown, digest: unknown): boolean {
  return isBoundedBoundary(boundary) && isUuid(transactionId) &&
    typeof digest === "string" && digest.length > 0 && digest.length <= MAX_DIGEST_LENGTH;
}

function boundedNonNegative(value: number, fallback: number): number {
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

function boundedPositiveInteger(value: number, fallback: number): number {
  return Number.isSafeInteger(value) && value > 0 ? value : fallback;
}

function utf8Length(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

function browserStorage(): IdempotencyStorage {
  try {
    if (typeof localStorage === "undefined") return new UnavailableIdempotencyStorage();
    return localStorage;
  } catch {
    return new UnavailableIdempotencyStorage();
  }
}

class UnavailableIdempotencyStorage implements IdempotencyStorage {
  getItem(): string | null { throw new Error("browser storage unavailable"); }
  setItem(): void { throw new Error("browser storage unavailable"); }
  removeItem(): void { throw new Error("browser storage unavailable"); }
}
